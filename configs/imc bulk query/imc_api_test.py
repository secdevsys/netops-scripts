#!/usr/bin/env python3
"""
Quick health check of HPE IMC REST API (imcrs) access.

Checks, in order:
  1. TCP connection to the IMC host/port
  2. HTTP/TLS response
  3. Digest authentication
  4. Device list API          (/imcrs/plat/res/device)
  5. Asset inventory API      (/imcrs/netasset/asset)
  6. Config backup API (ICC)  (/imcrs/icc/deviceCfg/<id>/currentRun)

Windows (PowerShell):
  py -m pip install requests
  $env:IMC_USER="admin"; $env:IMC_PASS="secret"
  py imc_api_test.py --url https://imc:8443
  py imc_api_test.py --url https://imc:8443 --insecure --dev-ip 10.20.1.1

Exit code 0 if all checks pass, 1 otherwise.
"""
import argparse
import os
import socket
import sys
import time
from urllib.parse import urlparse

try:
    import requests
    from requests.auth import HTTPDigestAuth
except ImportError:
    sys.exit("Missing 'requests'. Install it with:  py -m pip install requests")

results = []


def report(name, ok, detail=""):
    results.append(ok)
    tag = "PASS" if ok else "FAIL"
    if ok is None:
        tag = "SKIP"
    print(f"[{tag}] {name}" + (f"  - {detail}" if detail else ""))
    return ok


def as_list(x):
    return [x] if isinstance(x, dict) else (x or [])


def main():
    ap = argparse.ArgumentParser(description="Test HPE IMC REST API access")
    ap.add_argument("--url", required=True, help="e.g. http://imc:8080 or https://imc:8443")
    ap.add_argument("--user", default=os.getenv("IMC_USER"))
    ap.add_argument("--password", default=os.getenv("IMC_PASS"))
    ap.add_argument("--insecure", action="store_true", help="skip TLS certificate check")
    ap.add_argument("--dev-ip", help="device IP to use for the config test (default: first device)")
    ap.add_argument("--timeout", type=float, default=15)
    a = ap.parse_args()
    if not a.user or not a.password:
        sys.exit("Set --user/--password or IMC_USER/IMC_PASS.")

    base = a.url.rstrip("/")
    u = urlparse(base)
    host = u.hostname
    port = u.port or (443 if u.scheme == "https" else 80)
    print(f"Testing IMC at {base} as '{a.user}'\n")

    # 1. TCP
    try:
        t = time.time()
        socket.create_connection((host, port), timeout=a.timeout).close()
        report("TCP connection", True, f"{host}:{port} in {int((time.time() - t) * 1000)} ms")
    except OSError as e:
        report("TCP connection", False, f"{host}:{port} -> {e}. Check firewall/port/VPN.")
        return summary()

    s = requests.Session()
    s.headers["Accept"] = "application/json"
    s.verify = not a.insecure
    if a.insecure:
        requests.packages.urllib3.disable_warnings()

    # 2. HTTP / TLS (unauthenticated: 401 is the expected answer)
    try:
        r = s.get(base + "/imcrs", timeout=a.timeout)
        report("HTTP/TLS", True, f"HTTP {r.status_code}")
    except requests.exceptions.SSLError as e:
        report("HTTP/TLS", False, f"certificate error ({e.__class__.__name__}). Retry with --insecure")
        return summary()
    except requests.RequestException as e:
        hint = " (corporate proxy? set NO_PROXY to the IMC host)" if os.getenv("HTTPS_PROXY") or os.getenv("HTTP_PROXY") else ""
        report("HTTP/TLS", False, f"{e}{hint}")
        return summary()

    # 3. Auth + 4. device list
    s.auth = HTTPDigestAuth(a.user, a.password)
    try:
        r = s.get(base + "/imcrs/plat/res/device",
                  params={"resPrivilegeFilter": "false", "start": 0, "size": 5},
                  timeout=a.timeout)
    except requests.RequestException as e:
        report("Authentication", False, str(e))
        return summary()
    if r.status_code == 401:
        report("Authentication", False, "401 - wrong credentials, or operator lacks REST/API rights")
        return summary()
    if r.status_code == 403:
        report("Authentication", False, "403 - operator authenticated but not allowed to use the API")
        return summary()
    report("Authentication", True, "digest auth accepted")

    devices = []
    if r.ok:
        try:
            devices = as_list(r.json().get("device"))
            total = r.headers.get("Total") or "?"
            report("Device list API", True, f"{len(devices)} returned (sample), total header: {total}")
            for d in devices[:3]:
                print(f"         id={d.get('id')} ip={d.get('ip')} label={d.get('label')}")
        except ValueError:
            report("Device list API", False, "response is not JSON (wrong URL/port? web UI instead of API?)")
    else:
        report("Device list API", False, f"HTTP {r.status_code}: {r.text[:150]}")

    # 5. Assets
    try:
        r = s.get(base + "/imcrs/netasset/asset", params={"start": 0, "size": 5}, timeout=a.timeout)
        if r.ok:
            n = len(as_list(r.json().get("netAsset")))
            report("Asset API", True, f"{n} asset rows returned (sample)")
        else:
            report("Asset API", False, f"HTTP {r.status_code} - asset module/rights missing?")
    except (requests.RequestException, ValueError) as e:
        report("Asset API", False, str(e))

    # 6. Config (ICC)
    dev = None
    if a.dev_ip:
        try:
            r = s.get(base + "/imcrs/plat/res/device",
                      params={"resPrivilegeFilter": "false", "ip": a.dev_ip, "exact": "true"},
                      timeout=a.timeout)
            found = as_list(r.json().get("device")) if r.ok else []
            dev = found[0] if found else None
            if not dev:
                report("Config API (ICC)", False, f"device {a.dev_ip} not found in IMC")
        except (requests.RequestException, ValueError) as e:
            report("Config API (ICC)", False, str(e))
    elif devices:
        dev = devices[0]
    else:
        report("Config API (ICC)", None, "no device available to test")

    if dev:
        try:
            r = s.get(f"{base}/imcrs/icc/deviceCfg/{dev.get('id')}/currentRun", timeout=a.timeout * 2)
            content = r.json().get("content") if r.ok and r.text.strip() else None
            if content:
                report("Config API (ICC)", True,
                       f"{dev.get('ip')}: {len(content.splitlines())} lines of running config")
            elif r.ok:
                report("Config API (ICC)", False,
                       f"{dev.get('ip')}: no config stored - not backed up yet in ICC, try --dev-ip")
            else:
                report("Config API (ICC)", False,
                       f"HTTP {r.status_code} - ICC module not licensed or no rights")
        except (requests.RequestException, ValueError) as e:
            report("Config API (ICC)", False, str(e))

    return summary()


def summary():
    passed = sum(1 for x in results if x)
    failed = sum(1 for x in results if x is False)
    print(f"\n{passed} passed, {failed} failed")
    if failed == 0:
        print("IMC API access OK - you can run imc_bulk_export.py")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
