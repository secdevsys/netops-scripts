#!/usr/bin/env python3
"""
Bulk export from HPE/H3C Intelligent Management Center (IMC) via the imcrs REST API.
Built for large inventories: resumable, filterable, throttled, with retries.

Exports:
  - devices.csv   : device list (id, label, IP, type, status...)
  - assets.csv    : hardware inventory (netasset)
  - configs\\<ip>_<label>_running.cfg / _startup.cfg   (requires the ICC module)
  - failed.csv    : devices with no config returned (re-run with --retry-failed)

Windows (PowerShell):
  py -m pip install requests
  $env:IMC_USER="admin"; $env:IMC_PASS="secret"
  py imc_bulk_export.py --url https://imc:8443 --out C:\\imc_export --resume

Examples:
  --match ip=^10\\.10\\.            only devices whose IP matches the regex
  --match label=SW-PAR --match typeName=5130     several filters (all must match)
  --limit 200                      test on the first 200 devices
  --running-only                   skip startup configs (half the requests)
  --workers 4 --batch 100 --pause 30   gentler on the IMC server
  --resume                         skip devices whose config is already on disk
  --retry-failed                   only re-try devices listed in failed.csv
"""
import argparse
import csv
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
from requests.auth import HTTPDigestAuth
from urllib3.util.retry import Retry

PAGE = 500
HEADERS = {"Accept": "application/json", "Content-Type": "application/json"}


class IMC:
    def __init__(self, url, user, pwd, verify=True, timeout=90, workers=8):
        self.base = url.rstrip("/")
        self.s = requests.Session()
        self.s.auth = HTTPDigestAuth(user, pwd)  # IMC uses HTTP Digest auth
        self.s.headers.update(HEADERS)
        self.s.verify = verify
        retry = Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                      allowed_methods=("GET",))
        adapter = HTTPAdapter(max_retries=retry, pool_maxsize=max(10, workers * 2))
        self.s.mount("http://", adapter)
        self.s.mount("https://", adapter)
        self.timeout = timeout

    def get(self, path, **params):
        r = self.s.get(self.base + path, params=params, timeout=self.timeout)
        if r.status_code == 401:
            sys.exit("401 Unauthorized - check IMC_USER/IMC_PASS and REST API rights of the operator.")
        r.raise_for_status()
        return r.json() if r.text.strip() else {}

    def paged(self, path, key, label, **params):
        """IMC returns a dict instead of a list when there is exactly one item."""
        start = 0
        while True:
            data = self.get(path, start=start, size=PAGE, **params)
            items = data.get(key, [])
            if isinstance(items, dict):
                items = [items]
            yield from items
            start += len(items)
            print(f"\r  {label}: {start}", end="", flush=True)
            if len(items) < PAGE:
                print()
                break

    def devices(self):
        return list(self.paged("/imcrs/plat/res/device", "device", "devices",
                               resPrivilegeFilter="false"))

    def assets(self):
        return list(self.paged("/imcrs/netasset/asset", "netAsset", "asset rows"))

    def config(self, dev_id, kind):
        """kind: 'currentRun' or 'currentStart'. Needs the IMC ICC module."""
        try:
            return self.get(f"/imcrs/icc/deviceCfg/{dev_id}/{kind}").get("content")
        except (requests.RequestException, ValueError):
            return None


def write_csv(path, rows):
    if not rows:
        return 0
    fields = sorted({k for r in rows for k in r if not isinstance(r[k], (dict, list))})
    # utf-8-sig so Excel on Windows opens accents correctly
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    return len(rows)


def safe(s):
    return re.sub(r"[^\w.-]+", "_", str(s))[:80]


def cfg_path(outdir, dev, suffix):
    return os.path.join(outdir, f"{safe(dev.get('ip', ''))}_{safe(dev.get('label', ''))}_{suffix}.cfg")


def fetch_configs(imc, dev, outdir, kinds):
    saved = []
    for kind, suffix in kinds:
        cfg = imc.config(dev.get("id"), kind)
        if cfg:
            with open(cfg_path(outdir, dev, suffix), "w", encoding="utf-8", newline="") as f:
                f.write(cfg)
            saved.append(suffix)
    return dev, saved


def parse_matches(specs):
    out = []
    for m in specs:
        if "=" not in m:
            sys.exit(f"--match expects field=regex, got: {m}")
        k, v = m.split("=", 1)
        out.append((k, re.compile(v, re.I)))
    return out


def main():
    ap = argparse.ArgumentParser(description="Bulk export assets/configs from HPE IMC")
    ap.add_argument("--url", required=True, help="e.g. http://imc:8080 or https://imc:8443")
    ap.add_argument("--user", default=os.getenv("IMC_USER"))
    ap.add_argument("--password", default=os.getenv("IMC_PASS"))
    ap.add_argument("--out", default="imc_export")
    ap.add_argument("--match", action="append", default=[], help="field=regex filter on device list")
    ap.add_argument("--limit", type=int, default=0, help="process at most N devices")
    ap.add_argument("--workers", type=int, default=8, help="parallel config downloads")
    ap.add_argument("--batch", type=int, default=0, help="devices per batch (0 = no batching)")
    ap.add_argument("--pause", type=float, default=0, help="seconds to wait between batches")
    ap.add_argument("--resume", action="store_true", help="skip devices already saved")
    ap.add_argument("--retry-failed", action="store_true", help="only devices in failed.csv")
    ap.add_argument("--running-only", action="store_true")
    ap.add_argument("--no-configs", action="store_true")
    ap.add_argument("--no-assets", action="store_true")
    ap.add_argument("--insecure", action="store_true", help="skip TLS verification")
    a = ap.parse_args()
    if not a.user or not a.password:
        sys.exit("Set --user/--password or IMC_USER/IMC_PASS.")
    if a.insecure:
        requests.packages.urllib3.disable_warnings()

    imc = IMC(a.url, a.user, a.password, verify=not a.insecure, workers=a.workers)
    os.makedirs(a.out, exist_ok=True)
    failed_csv = os.path.join(a.out, "failed.csv")

    # ---- device list + filters
    devs = imc.devices()
    for field, rx in parse_matches(a.match):
        devs = [d for d in devs if rx.search(str(d.get(field, "")))]
    if a.retry_failed:
        if not os.path.exists(failed_csv):
            sys.exit("No failed.csv found.")
        with open(failed_csv, encoding="utf-8-sig") as f:
            ids = {row["id"] for row in csv.DictReader(f)}
        devs = [d for d in devs if str(d.get("id")) in ids]
    if a.limit:
        devs = devs[: a.limit]
    print(f"Selected devices: {len(devs)}")
    if not a.retry_failed:
        write_csv(os.path.join(a.out, "devices.csv"), devs)

    # ---- assets (one paged bulk call, not per device)
    if not a.no_assets and not a.retry_failed:
        ids = {str(d["id"]) for d in devs}
        assets = [x for x in imc.assets() if str(x.get("devId")) in ids]
        print(f"Asset rows written: {write_csv(os.path.join(a.out, 'assets.csv'), assets)}")

    if a.no_configs:
        return

    # ---- configs
    cdir = os.path.join(a.out, "configs")
    os.makedirs(cdir, exist_ok=True)
    kinds = [("currentRun", "running")] + ([] if a.running_only else [("currentStart", "startup")])

    todo = devs
    if a.resume:
        todo = [d for d in devs if not os.path.exists(cfg_path(cdir, d, "running"))]
        print(f"Resume: {len(devs) - len(todo)} already saved, {len(todo)} to do")

    batch = a.batch or len(todo) or 1
    ok, failed, done, t0 = 0, [], 0, time.time()
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        for i in range(0, len(todo), batch):
            chunk = todo[i:i + batch]
            for fut in as_completed([ex.submit(fetch_configs, imc, d, cdir, kinds) for d in chunk]):
                dev, saved = fut.result()
                done += 1
                if saved:
                    ok += 1
                else:
                    failed.append(dev)
                rate = done / max(time.time() - t0, 1e-6)
                eta = (len(todo) - done) / rate if rate else 0
                print(f"\r  configs {done}/{len(todo)}  ok={ok} failed={len(failed)}  "
                      f"ETA {int(eta // 60)}m{int(eta % 60):02d}s   ", end="", flush=True)
            if a.pause and i + batch < len(todo):
                time.sleep(a.pause)
    print()

    if failed:
        write_csv(failed_csv, failed)
        print(f"{len(failed)} devices without config -> {failed_csv} (use --retry-failed)")
    elif os.path.exists(failed_csv) and a.retry_failed:
        os.remove(failed_csv)
    print(f"Done: configs saved for {ok} devices in {int(time.time() - t0)}s")


if __name__ == "__main__":
    main()
