<#
.SYNOPSIS
    cp_topo.ps1 - Analyse d'un export HTML de configuration Check Point (cluster, couches/layers)
                  et generation de la topologie au format Visio (.vsdx) + SVG + rapport + CSV.

.DESCRIPTION
    Portage PowerShell de cp_topo.py. Aucune dependance : Windows PowerShell 5.1 ou PowerShell 7+.
      1. Extrait tous les tableaux HTML avec leur contexte (titres h1-h6, legendes, lignes-titres).
      2. Reconnait les colonnes par alias (Name / IPv4 Address / Net Mask / Topology / Source / Action ...).
      3. Reconstitue cluster, membres, interfaces, routes statiques et regles de chaque couche.
      4. Deduit la topologie (reseaux des VIP, synchro/gestion, routeurs de prochain saut).
      5. Produit : .vsdx (une page par cluster + page "Couches"), un .svg par page, un rapport .md,
         un CSV des regles et, en option, un JSON.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\cp_topo.ps1 .\export.html
.EXAMPLE
    .\cp_topo.ps1 .\dossier_export -Output .\topo.vsdx -Json .\modele.json
.EXAMPLE
    .\cp_topo.ps1 .\export.html -ShowTables          # affiche comment chaque tableau est interprete
.EXAMPLE
    .\cp_topo.ps1 .\export.html -Cluster CL-FW-PARIS  # force le nom de l'objet cluster
#>
param(
    [Parameter(Mandatory = $true, Position = 0)][string[]]$Path,
    [string]$Output,
    [string]$Report,
    [string]$Json,
    [string]$Cluster,
    [switch]$ShowTables
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = $Inv   # points decimaux dans le XML
$Utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
$Utf8Bom = New-Object System.Text.UTF8Encoding -ArgumentList $true
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# 0. Utilitaires
# ---------------------------------------------------------------------------

function Full([string]$p) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
function Collapse([string]$s) { if (-not $s) { return '' }; return ([regex]::Replace($s, '\s+', ' ')).Trim() }
function Norm([string]$s) {
    if (-not $s) { return '' }
    $s = $s.ToLower().Replace('_', ' ').Replace("`n", ' ')
    $s = [regex]::Replace($s, '[:*()\[\]]', ' ')
    return (Collapse $s)
}
function Esc($s) { if ($null -eq $s) { return '' }; return [System.Security.SecurityElement]::Escape([string]$s) }
function Fmt($v) {
    if ($v -is [double] -or $v -is [single] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]) {
        return ([double]$v).ToString('0.############', $Inv)
    }
    return [string]$v
}
function F1($v) { return ([double]$v).ToString('0.0', $Inv) }
function Write-Text([string]$p, [string]$content, $enc) { [System.IO.File]::WriteAllText($p, $content, $enc) }

function Read-Html([string]$p) {
    $bytes = [System.IO.File]::ReadAllBytes($p)
    try {
        $strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
        $txt = $strict.GetString($bytes)
    } catch {
        try { $txt = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes) }
        catch { $txt = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes) }
    }
    return $txt.TrimStart([char]0xFEFF)
}

# ---------------------------------------------------------------------------
# 1. Extraction des tableaux HTML (tokeniseur leger)
# ---------------------------------------------------------------------------

$HeadingTags = @('h1', 'h2', 'h3', 'h4', 'h5', 'h6')
$TagRe = [regex]@'
(?is)<!--.*?-->|<(script|style)\b.*?</\1\s*>|<!.*?>|<\?.*?>|<(/?)([a-z][a-z0-9]*)\b((?:[^>"']|"[^"]*"|'[^']*')*)>
'@

function Add-HtmlText($st, [string]$data, [bool]$isBr) {
    if (-not $isBr) { $data = $data -replace '[\r\n\t]+', ' ' }
    if ($st.stack.Count -gt 0) {
        $t = $st.stack[$st.stack.Count - 1]
        if ($null -ne $st.cbuf) { [void]$st.cbuf.Append($data) }
        elseif ($null -ne $t.cell) { [void]$t.cell.text.Append($data) }
    } elseif ($null -ne $st.hbuf) {
        [void]$st.hbuf.Append($data)
    } else {
        $s = Collapse $data
        if ($s) {
            $st.recent.Add($s)
            while ($st.recent.Count -gt 4) { $st.recent.RemoveAt(0) }
        }
    }
}

function Close-Cell($t) {
    if ($null -eq $t.cell) { return }
    $lines = @($t.cell.text.ToString() -split "`n" | ForEach-Object { Collapse $_ } | Where-Object { $_ })
    $t.row.Add(@{ text = ($lines -join "`n"); th = $t.cell.th; span = $t.cell.span })
    $t.cell = $null
}

function Close-Row($t) {
    Close-Cell $t
    if ($null -ne $t.row -and $t.row.Count -gt 0) { $t.rows.Add($t.row) }
    $t.row = $null
}

function Html-Start($st, [string]$tag, [string]$attrs) {
    $inTable = $st.stack.Count -gt 0
    $top = $null
    if ($inTable) { $top = $st.stack[$st.stack.Count - 1] }
    if ($HeadingTags -contains $tag) {
        if (-not $inTable) { $st.hbuf = New-Object System.Text.StringBuilder }
    } elseif ($tag -eq 'table') {
        $st.stack.Add(@{ rows = (New-Object System.Collections.Generic.List[object]); heading = $st.lastHeading;
                caption = ''; context = $st.recent.ToArray(); row = $null; cell = $null })
    } elseif ($tag -eq 'caption') {
        if ($inTable) { $st.cbuf = New-Object System.Text.StringBuilder }
    } elseif ($tag -eq 'tr') {
        if ($inTable) { Close-Row $top; $top.row = New-Object System.Collections.Generic.List[object] }
    } elseif ($tag -eq 'td' -or $tag -eq 'th') {
        if ($inTable) {
            Close-Cell $top
            if ($null -eq $top.row) { $top.row = New-Object System.Collections.Generic.List[object] }
            $span = 1
            $mm = [regex]::Match($attrs, 'colspan\s*=\s*["'']?(\d+)', 'IgnoreCase')
            if ($mm.Success) { $span = [math]::Max(1, [int]$mm.Groups[1].Value) }
            $top.cell = @{ text = (New-Object System.Text.StringBuilder); th = ($tag -eq 'th'); span = $span }
        }
    } elseif ($tag -eq 'br') {
        Add-HtmlText $st "`n" $true
    }
}

function Html-End($st, [string]$tag) {
    $inTable = $st.stack.Count -gt 0
    $top = $null
    if ($inTable) { $top = $st.stack[$st.stack.Count - 1] }
    if ($HeadingTags -contains $tag) {
        if ($null -ne $st.hbuf) {
            $txt = Collapse $st.hbuf.ToString()
            if ($txt) { $st.lastHeading = $txt; $st.recent.Clear() }
            $st.hbuf = $null
        }
    } elseif ($tag -eq 'caption') {
        if ($inTable -and $null -ne $st.cbuf) { $top.caption = Collapse $st.cbuf.ToString(); $st.cbuf = $null }
    } elseif ($tag -eq 'td' -or $tag -eq 'th') {
        if ($inTable) { Close-Cell $top }
    } elseif ($tag -eq 'tr') {
        if ($inTable) { Close-Row $top }
    } elseif ($tag -eq 'table') {
        if ($inTable) {
            Close-Row $top
            $st.stack.RemoveAt($st.stack.Count - 1)
            if ($top.rows.Count -gt 0) { $st.tables.Add($top) }
        }
    }
}

function Get-HtmlTables([string]$html) {
    $st = @{ tables = (New-Object System.Collections.Generic.List[object]);
             stack = (New-Object System.Collections.Generic.List[object]);
             lastHeading = ''; recent = (New-Object System.Collections.Generic.List[string]);
             hbuf = $null; cbuf = $null }
    $pos = 0
    foreach ($m in $TagRe.Matches($html)) {
        if ($m.Index -gt $pos) {
            Add-HtmlText $st ([System.Net.WebUtility]::HtmlDecode($html.Substring($pos, $m.Index - $pos))) $false
        }
        $pos = $m.Index + $m.Length
        if (-not $m.Groups[3].Success) { continue }
        $tag = $m.Groups[3].Value.ToLower()
        if ($m.Groups[2].Value -eq '/') { Html-End $st $tag } else { Html-Start $st $tag $m.Groups[4].Value }
    }
    if ($pos -lt $html.Length) { Add-HtmlText $st ([System.Net.WebUtility]::HtmlDecode($html.Substring($pos))) $false }
    while ($st.stack.Count -gt 0) { Html-End $st 'table' }
    return ,$st.tables
}

# ---------------------------------------------------------------------------
# 2. Reconnaissance des colonnes
# ---------------------------------------------------------------------------

$FieldAliases = [ordered]@{
    name      = @('name', 'interface', 'interface name', 'nom', 'if name', 'ifname', 'iface', 'nom interface',
                  "nom de l'interface", 'object name')
    ip        = @('ip address', 'ipv4 address', 'ip', 'ipv4', 'address', 'adresse ip', 'adresse', 'ip addr',
                  'ipv4 addr', 'ip/mask', 'ipv4 address/mask')
    mask      = @('net mask', 'netmask', 'mask', 'subnet mask', 'ipv4 net mask', 'ipv4 mask', 'network mask',
                  'mask length', 'ipv4 mask length', 'prefix', 'prefix length', 'masque',
                  'masque de sous-réseau', 'ipv4 network mask')
    topology  = @('topology', 'leads to', 'topologie', 'network type', 'type de réseau', 'network topology',
                  'interface topology', 'topology type')
    member    = @('member', 'cluster member', 'membre', 'gateway', 'device', 'member name', 'nom du membre',
                  'object', 'host')
    dest      = @('destination', 'destination network', 'dest', 'réseau destination', 'route', 'network destination')
    nexthop   = @('next hop', 'nexthop', 'next-hop', 'via', 'passerelle', 'gateway', 'next hop gateway', 'prochain saut')
    antispoof = @('anti spoofing', 'anti-spoofing', 'antispoofing')
    vlan      = @('vlan', 'vlan id', 'vlan tag')
}

$RuleAliases = [ordered]@{
    no      = @('no', 'no.', '#', 'rule', 'rule number', 'rule no', 'n°', 'numéro', 'num')
    rname   = @('name', 'rule name', 'nom')
    src     = @('source', 'src', 'sources')
    dst     = @('destination', 'dst', 'destinations')
    vpn     = @('vpn')
    svc     = @('services & applications', 'services and applications', 'services', 'service',
                'services & apps', 'applications')
    content = @('content', 'data', 'contenu')
    action  = @('action', 'actions')
    track   = @('track', 'log', 'logs', 'journalisation', 'suivi')
    install = @('install on', 'installé sur', 'install', 'target', 'targets')
    time    = @('time', 'heure')
    hits    = @('hits', 'hit count')
    comment = @('comments', 'comment', 'commentaires', 'commentaire')
}

function New-Map { return ,(New-Object 'System.Collections.Generic.Dictionary[int,string]') }

function Match-Field([string]$header) {
    $h = Norm $header
    if (-not $h) { return $null }
    foreach ($f in $FieldAliases.Keys) { if ($FieldAliases[$f] -contains $h) { return $f } }
    foreach ($f in $FieldAliases.Keys) {
        foreach ($al in $FieldAliases[$f]) {
            if ($al.Length -ge 6 -and ($h.StartsWith($al) -or $h.Contains($al))) { return $f }
        }
    }
    return $null
}

function Expand-Row($row) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($c in $row) {
        $out.Add([string]$c.text)
        for ($i = 1; $i -lt $c.span; $i++) { $out.Add('') }
    }
    return ,$out.ToArray()
}

function Get-Width($rows) {
    $w = 0
    foreach ($r in $rows) { $n = (Expand-Row $r).Count; if ($n -gt $w) { $w = $n } }
    return $w
}

function Split-Subtables($t) {
    # Decoupe un tableau sur les lignes-titres (une seule cellule non vide)
    $width = Get-Width $t.rows
    $parts = New-Object System.Collections.Generic.List[object]
    $cur = New-Object System.Collections.Generic.List[object]
    $sub = ''
    foreach ($r in $t.rows) {
        $cells = Expand-Row $r
        $nonEmpty = @($cells | Where-Object { $_.Trim() })
        if ($width -ge 2 -and $nonEmpty.Count -eq 1 -and ($r.Count -eq 1 -or $r[0].span -gt 1)) {
            if ($cur.Count -gt 0) { $parts.Add(@{ sub = $sub; rows = $cur }) }
            $sub = $nonEmpty[0]
            $cur = New-Object System.Collections.Generic.List[object]
        } else {
            $cur.Add($r)
        }
    }
    if ($cur.Count -gt 0) { $parts.Add(@{ sub = $sub; rows = $cur }) }
    return ,$parts
}

function Analyse-Header($rows) {
    $n = [math]::Min(3, $rows.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $r = $rows[$i]
        $cells = Expand-Row $r
        $map = New-Map
        for ($j = 0; $j -lt $cells.Count; $j++) {
            $f = Match-Field $cells[$j]
            if ($f -and -not $map.ContainsValue($f)) { $map[$j] = $f }
        }
        $allTh = $true
        foreach ($c in $r) { if (-not $c.th) { $allTh = $false } }
        if ($map.Count -ge 2 -or ($allTh -and $map.Count -gt 0)) {
            if ($map.ContainsValue('dest')) {   # 'gateway' = prochain saut dans un tableau de routes
                foreach ($j in @($map.Keys)) {
                    $nh = Norm $cells[$j]
                    if ($map[$j] -eq 'member' -and ($nh -eq 'gateway' -or $nh -eq 'passerelle')) { $map[$j] = 'nexthop' }
                }
            }
            return @{ map = $map; start = $i + 1 }
        }
    }
    return @{ map = (New-Map); start = 0 }
}

function Analyse-RuleHeader($rows) {
    $n = [math]::Min(3, $rows.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $cells = Expand-Row $rows[$i]
        $map = New-Map
        for ($j = 0; $j -lt $cells.Count; $j++) {
            $h = Norm $cells[$j]
            if (-not $h) { continue }
            foreach ($f in $RuleAliases.Keys) {
                $al = $RuleAliases[$f]
                $hit = ($al -contains $h)
                if (-not $hit) { foreach ($a in $al) { if ($a.Length -ge 6 -and $h.StartsWith($a)) { $hit = $true; break } } }
                if ($hit -and -not $map.ContainsValue($f)) { $map[$j] = $f; break }
            }
        }
        $vals = @($map.Values)
        if ($vals -contains 'src' -and $vals -contains 'dst' -and $vals -contains 'action') {
            return @{ map = $map; start = $i + 1 }
        }
    }
    return @{ map = (New-Map); start = 0 }
}

function Layer-FromText([string]$t) {
    # 'Network Layer', 'Layer: Appli', 'Standard - Network layer' -> nom de couche
    if (-not $t -or -not [regex]::IsMatch($t, '\b(layer|layers|couche)\b', 'IgnoreCase')) { return $null }
    $parts = [regex]::Split($t, '\s+[-–|]\s+')
    $s = $parts[$parts.Count - 1]
    $s = [regex]::Replace($s, '\b(access control|policy|politique|ordered|inline)\b', ' ', 'IgnoreCase')
    $s = [regex]::Replace($s, '\b(layers?|couches?)\b\s*[:\-]?', ' ', 'IgnoreCase')
    $s = Collapse $s
    $s = [regex]::Replace($s, '^(de la|du|de|of|:)\s+', '', 'IgnoreCase')
    $s = $s.Trim([char[]]@(' ', ':', '-', '"', "'"))
    if ($s) { return $s }
    return $null
}

# ---------------------------------------------------------------------------
# 3. Valeurs reseau (IPv4)
# ---------------------------------------------------------------------------

$IpRe = '\b(\d{1,3}(?:\.\d{1,3}){3})(?:\s*/\s*(\d{1,2}))?\b'

function Test-Ip([string]$ip) {
    $o = $ip.Split('.')
    if ($o.Count -ne 4) { return $false }
    foreach ($oc in $o) { if ([int]$oc -gt 255) { return $false } }
    return $true
}
function IpToInt([string]$ip) {
    $o = $ip.Split('.')
    return ([int64]$o[0] * 16777216) + ([int64]$o[1] * 65536) + ([int64]$o[2] * 256) + [int64]$o[3]
}
function IntToIp([int64]$n) {
    $a = [int64]([math]::Floor($n / 16777216)) % 256
    $b = [int64]([math]::Floor($n / 65536)) % 256
    $c = [int64]([math]::Floor($n / 256)) % 256
    $d = $n % 256
    return ('{0}.{1}.{2}.{3}' -f $a, $b, $c, $d)
}
function PrefixMask([int]$p) {
    if ($p -le 0) { return [int64]0 }
    return [int64]4294967295 - ([int64][math]::Pow(2, 32 - $p) - 1)
}
function Mask-ToPrefix([string]$mask) {
    $n = IpToInt $mask
    for ($p = 0; $p -le 32; $p++) {
        $pm = PrefixMask $p
        if ($pm -eq $n) { return $p }
        if (([int64]4294967295 - $pm) -eq $n) { return $p }   # masque inverse (wildcard)
    }
    return $null
}
function Parse-Prefix([string]$txt) {
    if (-not $txt) { return $null }
    $m = [regex]::Match($txt, $IpRe)
    if ($m.Success -and (Test-Ip $m.Groups[1].Value)) {
        $p = Mask-ToPrefix $m.Groups[1].Value
        if ($null -ne $p) { return $p }
    }
    $m = [regex]::Match($txt, '/?\b(\d{1,2})\b')
    if ($m.Success -and [int]$m.Groups[1].Value -le 32) { return [int]$m.Groups[1].Value }
    return $null
}
function Parse-Ip([string]$ipTxt, [string]$maskTxt) {
    # '10.0.0.1/24' ou ('10.0.0.1', '255.255.255.0') -> @{ip; prefix}
    $m = [regex]::Match($ipTxt, $IpRe)
    if (-not $m.Success -or -not (Test-Ip $m.Groups[1].Value)) { return @{ ip = $null; prefix = $null } }
    $pfx = $null
    if ($m.Groups[2].Success -and [int]$m.Groups[2].Value -le 32) { $pfx = [int]$m.Groups[2].Value }
    else { $pfx = Parse-Prefix $maskTxt }
    return @{ ip = $m.Groups[1].Value; prefix = $pfx }
}
function Net-Of($i) {
    if ($i.ip -and $null -ne $i.prefix) {
        $n = (IpToInt $i.ip) -band (PrefixMask $i.prefix)
        return ('{0}/{1}' -f (IntToIp $n), $i.prefix)
    }
    return $null
}
function In-Net([string]$ip, [string]$net) {
    $parts = $net.Split('/')
    $pm = PrefixMask ([int]$parts[1])
    return (((IpToInt $ip) -band $pm) -eq ((IpToInt $parts[0]) -band $pm))
}
function Short-Ip([string]$ip) { return '.' + $ip.Split('.')[-1] }

function Classify-Topology([string]$txt, [string]$ifname) {
    $t = (Norm $txt) + ' ' + (Norm $ifname)
    if ($t -match 'external|externe|internet') { return 'external' }
    if ($t.Contains('dmz')) { return 'dmz' }
    if ($t.Contains('sync')) { return 'sync' }
    if ($t -match 'mgmt|management|gestion') { return 'mgmt' }
    if ($t -match 'internal|interne|this network|specific|leads to|not defined') { return 'internal' }
    return 'unknown'
}

# ---------------------------------------------------------------------------
# 4. Reconstitution des objets
# ---------------------------------------------------------------------------

$ObjHeadRe = '(cluster\s*member|cluster|membre|member|gateway|passerelle|gw|firewall)(?:\s+object)?\s*[:=]\s*["'']?([A-Za-z0-9_.\-]+)'
$ObjParenRe = '^([A-Za-z0-9_.\-]+)\s*\(([^)]*)\)'
$KvNameKeys = @('name', 'object name', 'nom', "nom de l'objet")
$KvTypeKeys = @('type', 'object type', 'class', 'classe', 'object class')
$KvMembersKeys = @('cluster members', 'members', 'membres', 'membres du cluster')
$KvIpKeys = @('ipv4 address', 'ip address', 'main ip', 'adresse ip', 'ip')

function Kind-FromText([string]$t) {
    $t = Norm $t
    if ($t.Contains('member') -or $t.Contains('membre')) { return 'member' }
    if ($t.Contains('cluster')) { return 'cluster' }
    foreach ($k in @('gateway', 'passerelle', 'gw', 'firewall', 'simple')) { if ($t.Contains($k)) { return 'gateway' } }
    return $null
}

function Parse-ObjHeading([string]$text) {
    if (-not $text) { return $null }
    $m = [regex]::Match($text, $ObjHeadRe, 'IgnoreCase')
    if ($m.Success) { return @{ kind = (Kind-FromText $m.Groups[1].Value); name = $m.Groups[2].Value } }
    $m = [regex]::Match($text.Trim(), $ObjParenRe)
    if ($m.Success -and (Kind-FromText $m.Groups[2].Value)) {
        return @{ kind = (Kind-FromText $m.Groups[2].Value); name = $m.Groups[1].Value }
    }
    return $null
}

function New-Model {
    return @{ objects = [ordered]@{}; routes = (New-Object System.Collections.Generic.List[object]);
              rules = (New-Object System.Collections.Generic.List[object]); layers = [ordered]@{};
              log = (New-Object System.Collections.Generic.List[object]) }
}

function Get-Obj($model, [string]$name) {
    if (-not $model.objects.Contains($name)) {
        $model.objects[$name] = [ordered]@{
            name = $name; kind = $null; type = ''; ip = $null; cluster = $null
            members = (New-Object System.Collections.Generic.List[string])
            interfaces = (New-Object System.Collections.Generic.List[object])
            props = [ordered]@{}
        }
    }
    return $model.objects[$name]
}

function Add-Iface($model, [string]$owner, $iface) {
    $o = Get-Obj $model $owner
    foreach ($ex in $o.interfaces) {
        if ($ex.name -eq $iface.name -and $ex.ip -eq $iface.ip) {
            foreach ($k in @($iface.Keys)) { if ($iface[$k] -and -not $ex[$k]) { $ex[$k] = $iface[$k] } }
            return
        }
    }
    $o.interfaces.Add($iface)
}

function Find-KnownName([string]$text, $model) {
    if (-not $text) { return $null }
    $best = $null
    foreach ($n in @($model.objects.Keys)) {
        if ($n -and [regex]::IsMatch($text, '(?<![\w.\-])' + [regex]::Escape($n) + '(?![\w.\-])')) {
            if ($null -eq $best -or $n.Length -gt $best.Length) { $best = $n }
        }
    }
    return $best
}

function First-Kv($kv, $keys) {
    foreach ($k in $keys) { if ($kv.Contains($k) -and $kv[$k]) { return $kv[$k] } }
    return $null
}

function Get-Rec($map, $r) {
    $rec = @{}
    foreach ($j in $map.Keys) { if ($j -lt $r.Count) { $rec[$map[$j]] = [string]$r[$j] } else { $rec[$map[$j]] = '' } }
    return $rec
}

function Map-ToString($map) { return (@($map.Keys | ForEach-Object { "$_=$($map[$_])" }) -join ', ') }

function Build-Model($files) {
    $model = New-Model
    $curObj = $null; $curCluster = $null; $curLayer = $null

    foreach ($file in $files) {
        $src = Split-Path $file -Leaf
        $tables = Get-HtmlTables (Read-Html $file)
        foreach ($t in $tables) {
            $prevRmap = $null; $prevRw = -1
            foreach ($part in (Split-Subtables $t)) {
                $sub = [string]$part.sub; $rows = $part.rows
                $cap = [string]$t.caption; $head = [string]$t.heading
                $ctx = ((@($head, $cap, $sub) + @($t.context)) | Where-Object { $_ }) -join ' | '

                # Objet nomme par un titre ("Cluster: CL-FW", "fw01 (Cluster Member)")
                foreach ($cc in @($sub, $cap, $head)) {
                    $oh = Parse-ObjHeading $cc
                    if ($null -ne $oh) {
                        $o = Get-Obj $model $oh.name
                        if (-not $o.kind) { $o.kind = $oh.kind }
                        if ($oh.kind -eq 'member' -and $curCluster -and -not $o.cluster) { $o.cluster = $curCluster }
                        $curObj = $oh.name
                        if ($oh.kind -eq 'cluster') { $curCluster = $oh.name }
                        break
                    }
                }

                $width = Get-Width $rows

                # Couche (layer) courante
                foreach ($cc in @($sub, $cap, $head)) {
                    $ln = Layer-FromText $cc
                    if ($ln) { $curLayer = $ln; break }
                }

                # --- Regles de politique (une ou plusieurs couches) ------------
                $rh = Analyse-RuleHeader $rows
                $rmap = $rh.map; $rstart = $rh.start
                if ($rmap.Count -eq 0 -and $null -ne $prevRmap -and $width -eq $prevRw) { $rmap = $prevRmap; $rstart = 0 }
                if ($rmap.Count -gt 0) {
                    $prevRmap = $rmap; $prevRw = $width
                    if ($curLayer) { $layer = $curLayer } elseif ($head) { $layer = $head } else { $layer = 'Couche 1' }
                    $section = ''
                    if ($sub -and -not (Layer-FromText $sub)) { $section = $sub }
                    if (-not $model.layers.Contains($layer)) {
                        $model.layers[$layer] = [ordered]@{ name = $layer; rules = 0;
                            sources = (New-Object System.Collections.Generic.List[string]) }
                    }
                    $lay = $model.layers[$layer]
                    if (-not $lay.sources.Contains($src)) { $lay.sources.Add($src) }
                    for ($i = $rstart; $i -lt $rows.Count; $i++) {
                        $cells = Expand-Row $rows[$i]
                        $rec = [ordered]@{}
                        foreach ($j in $rmap.Keys) {
                            if ($j -lt $cells.Count) { $rec[$rmap[$j]] = $cells[$j].Trim() } else { $rec[$rmap[$j]] = '' }
                        }
                        if (-not ($rec['src'] -or $rec['dst'] -or $rec['action'])) { continue }
                        $rec['layer'] = $layer; $rec['section'] = $section
                        $rec['disabled'] = [regex]::IsMatch(($cells -join ' '), 'disabled|désactiv', 'IgnoreCase')
                        $model.rules.Add($rec)
                        $lay['rules'] = $lay['rules'] + 1
                    }
                    $hdr = @(); if ($rstart -gt 0) { $hdr = Expand-Row $rows[$rstart - 1] }
                    $model.log.Add([ordered]@{ source = $src; contexte = $ctx.Substring(0, [math]::Min(120, $ctx.Length));
                        colonnes = (Map-ToString $rmap); lignes = $rows.Count; interpretation = "règles [$layer]"; entete = $hdr })
                    continue
                }

                $hd = Analyse-Header $rows
                $map = $hd.map; $start = $hd.start
                $fields = @($map.Values)
                $data = New-Object System.Collections.Generic.List[object]
                for ($i = $start; $i -lt $rows.Count; $i++) { $data.Add((Expand-Row $rows[$i])) }
                $kindTab = '?'
                $hasSlashIp = $false
                foreach ($r in $data) { foreach ($j in $map.Keys) { if ($map[$j] -eq 'ip' -and $j -lt $r.Count -and $r[$j].Contains('/')) { $hasSlashIp = $true } } }

                # --- Tableau cle / valeur ----------------------------------------
                if ($width -eq 2 -and -not (($fields -contains 'ip' -and $fields -contains 'mask') -or $fields -contains 'dest')) {
                    $kindTab = 'clé/valeur'
                    $kv = [ordered]@{}
                    foreach ($r in $rows) {
                        $rr = Expand-Row $r
                        if ($rr.Count -ge 2 -and $rr[0].Trim()) { $kv[[string](Norm $rr[0])] = $rr[1].Trim() }
                    }
                    $name = First-Kv $kv $KvNameKeys
                    if ($name) { $name = ($name.Trim() -split '\s+')[0] } elseif ($curObj) { $name = $curObj }
                    if ($name) {
                        $o = Get-Obj $model $name
                        foreach ($k in @($kv.Keys)) { $o.props[$k] = $kv[$k] }
                        $typ = First-Kv $kv $KvTypeKeys
                        if ($typ) { $o.type = $typ; $kk = Kind-FromText $typ; if ($kk) { $o.kind = $kk } }
                        $ipv = First-Kv $kv $KvIpKeys
                        if ($ipv -and -not $o.ip) { $o.ip = (Parse-Ip $ipv '').ip }
                        $mem = First-Kv $kv $KvMembersKeys
                        if ($mem) {
                            if (-not $o.kind) { $o.kind = 'cluster' }
                            foreach ($mn in ($mem -split '[,;\n]+|\s{2,}')) {
                                $mn = ($mn.Trim() -split ' ')[0]
                                if ($mn -and -not $o.members.Contains($mn)) { $o.members.Add($mn) }
                            }
                        }
                        $curObj = $name
                        if ($o.kind -eq 'cluster') { $curCluster = $name }
                        elseif ($o.kind -eq 'member' -and $curCluster -and -not $o.cluster) { $o.cluster = $curCluster }
                    }
                }
                # --- Routes statiques --------------------------------------------
                elseif ($fields -contains 'dest' -and ($fields -contains 'nexthop' -or $fields -contains 'name')) {
                    $kindTab = 'routes'
                    $owner = Find-KnownName $ctx $model
                    if (-not $owner) { $owner = $curObj }
                    if (-not $owner) { $owner = '?' }
                    foreach ($r in $data) {
                        $rec = Get-Rec $map $r
                        $d = Parse-Ip $rec['dest'] $rec['mask']
                        if ((Norm $rec['dest']).Contains('default')) { $d = @{ ip = '0.0.0.0'; prefix = 0 } }
                        $nh = (Parse-Ip $rec['nexthop'] '').ip
                        if ($null -ne $d.ip -and $nh) {
                            $pf = 32; if ($null -ne $d.prefix) { $pf = $d.prefix }
                            $model.routes.Add(@{ owner = $owner; dest = $d.ip; prefix = $pf; nexthop = $nh })
                        }
                    }
                }
                # --- Liste des membres (nom [+ IP], sans masque) -----------------
                elseif (($fields -contains 'name' -or $fields -contains 'member') -and -not ($fields -contains 'mask') -and
                        -not $hasSlashIp -and [regex]::IsMatch("$sub $cap $head", 'member|membre', 'IgnoreCase')) {
                    $kindTab = 'membres'
                    $cl = $curCluster; if (-not $cl) { $cl = $curObj }
                    $col = 0; foreach ($j in $map.Keys) { if ($map[$j] -eq 'member' -or $map[$j] -eq 'name') { $col = $j; break } }
                    $ipcol = -1; foreach ($j in $map.Keys) { if ($map[$j] -eq 'ip') { $ipcol = $j; break } }
                    foreach ($r in $data) {
                        $mname = ''
                        if ($col -lt $r.Count) { $mname = ([string]$r[$col]).Trim() }
                        $mname = ($mname -split '\s+')[0]
                        if (-not $mname) { continue }
                        $mo = Get-Obj $model $mname
                        $mo.kind = 'member'
                        if ($cl) {
                            $mo.cluster = $cl
                            $co = Get-Obj $model $cl
                            if (-not $co.kind) { $co.kind = 'cluster' }
                            if (-not $co.members.Contains($mname)) { $co.members.Add($mname) }
                        }
                        if ($ipcol -ge 0 -and $ipcol -lt $r.Count -and -not $mo.ip) { $mo.ip = (Parse-Ip $r[$ipcol] '').ip }
                    }
                }
                # --- Tableau d'interfaces ----------------------------------------
                elseif ($fields -contains 'ip' -and ($fields -contains 'name' -or $fields -contains 'mask')) {
                    $kindTab = 'interfaces'
                    $ctxOwner = $null
                    foreach ($cc in @($sub, $cap, $head)) { $ctxOwner = Find-KnownName $cc $model; if ($ctxOwner) { break } }
                    $ownerDefault = $ctxOwner
                    if (-not $ownerDefault) { $ownerDefault = $curObj }
                    foreach ($r in $data) {
                        $rec = Get-Rec $map $r
                        $p = Parse-Ip $rec['ip'] $rec['mask']
                        if (-not $p.ip) { continue }
                        $owner = (([string]$rec['member']).Trim() -split '\s+')[0]
                        if (-not $owner) { $owner = $ownerDefault }
                        if (-not $owner) { $owner = 'FIREWALL' }
                        if ($rec['member'] -and $curCluster -and $owner -ne $curCluster) {
                            $mo = Get-Obj $model $owner
                            if (-not $mo.kind) { $mo.kind = 'member' }
                            if (-not $mo.cluster) { $mo.cluster = $curCluster }
                        }
                        $nm = ([string]$rec['name'] -split "`n")[0]
                        if (-not $nm) { $nm = '?' }
                        Add-Iface $model $owner ([ordered]@{
                            name = $nm; ip = $p.ip; prefix = $p.prefix
                            topology_raw = [string]$rec['topology']
                            topology = (Classify-Topology $rec['topology'] $rec['name'])
                            vlan = [string]$rec['vlan'] })
                    }
                }

                $hdr = @()
                if ($start -gt 0) { $hdr = Expand-Row $rows[$start - 1] } elseif ($rows.Count -gt 0) { $hdr = Expand-Row $rows[0] }
                $model.log.Add([ordered]@{ source = $src; contexte = $ctx.Substring(0, [math]::Min(120, $ctx.Length));
                    colonnes = (Map-ToString $map); lignes = $rows.Count; interpretation = $kindTab; entete = $hdr })
            }
        }
    }
    return $model
}

# ---------------------------------------------------------------------------
# 5. Deduction de la topologie
# ---------------------------------------------------------------------------

function Get-Net($networks, [string]$n, [string]$cls) {
    if (-not $networks.Contains($n)) {
        $networks[$n] = [ordered]@{ network = $n; role = $cls;
            vip = (New-Object System.Collections.Generic.List[object]);
            members = (New-Object System.Collections.Generic.List[object]); routers = [ordered]@{} }
    } elseif ($networks[$n].role -eq 'unknown' -and $cls -ne 'unknown') {
        $networks[$n].role = $cls
    }
    return $networks[$n]
}

function Run-Checks($cl, $members, $networks) {
    $out = New-Object System.Collections.Generic.List[string]
    $seen = [ordered]@{}
    foreach ($o in (@($cl) + @($members))) {
        foreach ($i in $o.interfaces) {
            if ($null -eq $i.prefix) { $out.Add(('Masque inconnu : {0} {1} ({2})' -f $o.name, $i.name, $i.ip)) }
            if (-not $seen.Contains($i.ip)) { $seen[$i.ip] = New-Object System.Collections.Generic.List[string] }
            $seen[$i.ip].Add(('{0}/{1}' -f $o.name, $i.name))
        }
    }
    foreach ($k in @($seen.Keys)) {
        if ($seen[$k].Count -gt 1) { $out.Add(('Adresse IP en double {0} : {1}' -f $k, ($seen[$k] -join ', '))) }
    }
    foreach ($n in $networks.Values) {
        if ($n.vip.Count -gt 0) {
            $have = @($n.members | ForEach-Object { $_.member })
            foreach ($m in $members) {
                if ($have -notcontains $m.name) {
                    $out.Add(("Le membre {0} n'a pas d'interface dans le réseau {1} (VIP {2})" -f $m.name, $n.network, $n.vip[0].ip))
                }
            }
            foreach ($xm in $n.members) {
                if ($xm.iface -ne $n.vip[0].iface -and $n.vip[0].iface -ne '?') {
                    $out.Add(('Nom d''interface différent sur {0} : VIP {1} / {2} {3}' -f $n.network, $n.vip[0].iface, $xm.member, $xm.iface))
                }
            }
        }
        if ($n.role -eq 'unknown' -and $n.vip.Count -gt 0) { $out.Add(('Topologie non renseignée pour {0}' -f $n.network)) }
    }
    if ($members.Count -lt 2) { $out.Add(('Moins de 2 membres détectés pour le cluster {0}' -f $cl.name)) }
    return ,$out
}

function Build-Topology($model, [string]$forced) {
    $objs = $model.objects
    $clusters = New-Object System.Collections.Generic.List[object]
    if ($forced) { (Get-Obj $model $forced).kind = 'cluster' }
    foreach ($o in @($objs.Values)) {
        $nt = Norm $o.type
        if ($o.kind -eq 'cluster' -or ($nt.Contains('cluster') -and -not $nt.Contains('member'))) { $clusters.Add($o) }
    }
    if ($clusters.Count -eq 0) {   # pas de cluster explicite : objet avec interfaces
        foreach ($o in @($objs.Values)) { if ($o.interfaces.Count -gt 0 -and $o.kind -ne 'member') { $clusters.Add($o); break } }
        if ($clusters.Count -eq 0) { foreach ($o in @($objs.Values)) { if ($o.interfaces.Count -gt 0) { $clusters.Add($o); break } } }
    }

    $topos = New-Object System.Collections.Generic.List[object]
    foreach ($cl in $clusters) {
        $members = New-Object System.Collections.Generic.List[object]
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($n in @($cl.members)) { $members.Add((Get-Obj $model $n)); $names.Add($n) }
        foreach ($o in @($objs.Values)) {
            if ($o.kind -eq 'member' -and -not $names.Contains($o.name)) {
                if ($o.cluster -eq $cl.name -or $null -eq $o.cluster -or $clusters.Count -eq 1) { $members.Add($o); $names.Add($o.name) }
            }
        }
        $vipNets = @(foreach ($i in $cl.interfaces) { $nn = Net-Of $i; if ($nn) { $nn } })
        if ($members.Count -eq 0) {   # passerelles partageant les reseaux des VIP
            foreach ($o in @($objs.Values)) {
                if ([object]::ReferenceEquals($o, $cl) -or $o.interfaces.Count -eq 0 -or $o.kind -eq 'cluster') { continue }
                $hit = $false
                foreach ($i in $o.interfaces) { $nn = Net-Of $i; if ($nn -and $vipNets -contains $nn) { $hit = $true } }
                if ($hit) { $members.Add($o) }
            }
        }

        $networks = [ordered]@{}
        foreach ($i in $cl.interfaces) {
            $n = Net-Of $i
            if (-not $n) { continue }
            (Get-Net $networks $n $i.topology).vip.Add(@{ iface = $i.name; ip = $i.ip })
        }
        foreach ($m in $members) {
            foreach ($i in $m.interfaces) {
                $n = Net-Of $i
                if (-not $n) { continue }
                $cls = $i.topology
                if (-not $networks.Contains($n) -and ($cls -eq 'unknown' -or $cls -eq 'internal')) {
                    if ($i.name.ToLower().Contains('mgmt')) { $cls = 'mgmt' } else { $cls = 'sync' }
                }
                (Get-Net $networks $n $cls).members.Add(@{ member = $m.name; iface = $i.name; ip = $i.ip })
            }
        }

        # Routes statiques -> routeurs de prochain saut
        $owners = @($cl.name, '?') + @($members | ForEach-Object { $_.name })
        foreach ($r in $model.routes) {
            if ($owners -notcontains $r.owner) { continue }
            foreach ($info in @($networks.Values)) {
                if (In-Net $r.nexthop $info.network) {
                    if ($r.prefix -eq 0) { $dest = 'défaut' } else { $dest = '{0}/{1}' -f $r.dest, $r.prefix }
                    if (-not $info.routers.Contains($r.nexthop)) { $info.routers[$r.nexthop] = New-Object System.Collections.Generic.List[string] }
                    $lst = $info.routers[$r.nexthop]
                    if (-not $lst.Contains($dest)) { $lst.Add($dest) }
                }
            }
        }

        $checks = Run-Checks $cl $members $networks
        $topos.Add(@{ cluster = $cl; members = $members; networks = @($networks.Values); checks = $checks })
    }
    return ,$topos
}

# ---------------------------------------------------------------------------
# 6. Ecriture Visio (.vsdx) - Open Packaging Conventions, sans dependance
# ---------------------------------------------------------------------------

$NS = 'http://schemas.microsoft.com/office/visio/2012/main'
$NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$RELV = 'http://schemas.microsoft.com/visio/2010/relationships/'

$STYLE = @{
    external = @('#F8CECC', '#B85450'); internal = @('#D5E8D4', '#82B366'); dmz = @('#FFE6CC', '#D79B00')
    sync = @('#E6E6E6', '#666666'); mgmt = @('#E1D5E7', '#9673A6'); unknown = @('#F5F5F5', '#999999')
    cluster = @('#DAE8FC', '#6C8EBF'); member = @('#FFFFFF', '#1F4E79'); router = @('#FFF2CC', '#B8860B')
}
$LABELS = @{ external = 'Externe'; internal = 'Interne'; dmz = 'DMZ'; sync = 'Synchro'; mgmt = 'Gestion'; unknown = 'Topologie ?' }
function Get-Style([string]$r) { if ($STYLE.ContainsKey($r)) { return ,$STYLE[$r] }; return ,$STYLE['unknown'] }
function Label-Of([string]$r) { if ($LABELS.ContainsKey($r)) { return $LABELS[$r] }; return $r }

function VCell([string]$n, $v, [string]$f, [string]$u) {
    $extra = ''
    if ($u) { $extra += ' U="' + $u + '"' }
    if ($f) { $extra += ' F="' + (Esc $f) + '"' }
    return '<Cell N="' + $n + '" V="' + (Esc (Fmt $v)) + '"' + $extra + '/>'
}

function New-VPage([string]$name) {
    return @{ name = $name; back = (New-Object System.Collections.Generic.List[string]);
              front = (New-Object System.Collections.Generic.List[string]);
              connects = (New-Object System.Collections.Generic.List[string]);
              svgBack = (New-Object System.Collections.Generic.List[object]);
              svgFront = (New-Object System.Collections.Generic.List[object]);
              nextId = 1; width = 11.0; height = 8.5 }
}

function Add-Box($pg, [double]$cx, [double]$cy, [double]$w, [double]$h, [string]$text, [string]$fill, [string]$line,
                 [double]$size = 9, [bool]$bold = $false, [double]$rounding = 0.08, [bool]$ellipse = $false,
                 [int]$valign = 1, [double]$lw = 1.0, [bool]$back = $false) {
    $sid = $pg.nextId
    $pg.nextId = $sid + 1
    $rnd = $rounding; if ($ellipse) { $rnd = 0 }
    $cells = @(
        (VCell 'PinX' $cx), (VCell 'PinY' $cy), (VCell 'Width' $w), (VCell 'Height' $h),
        (VCell 'LocPinX' ($w / 2) 'Width*0.5'), (VCell 'LocPinY' ($h / 2) 'Height*0.5'),
        (VCell 'Angle' 0), (VCell 'FillForegnd' $fill), (VCell 'FillPattern' 1),
        (VCell 'LineColor' $line), (VCell 'LineWeight' ($lw / 72.0)), (VCell 'LinePattern' 1),
        (VCell 'Rounding' $rnd), (VCell 'VerticalAlign' $valign),
        (VCell 'TopMargin' 0.05), (VCell 'BottomMargin' 0.05), (VCell 'LeftMargin' 0.06), (VCell 'RightMargin' 0.06)
    ) -join ''
    $flags = (VCell 'NoFill' 0) + (VCell 'NoLine' 0) + (VCell 'NoShow' 0)
    if ($ellipse) {
        $geom = '<Section N="Geometry" IX="0">' + $flags + '<Row T="Ellipse" IX="1">' +
                (VCell 'X' ($w / 2) 'Width*0.5') + (VCell 'Y' ($h / 2) 'Height*0.5') +
                (VCell 'A' $w 'Width*1') + (VCell 'B' ($h / 2) 'Height*0.5') +
                (VCell 'C' ($w / 2) 'Width*0.5') + (VCell 'D' $h 'Height*1') + '</Row></Section>'
    } else {
        $geom = '<Section N="Geometry" IX="0">' + $flags +
                '<Row T="RelMoveTo" IX="1">' + (VCell 'X' 0) + (VCell 'Y' 0) + '</Row>'
        $pts = @(@(1, 0), @(1, 1), @(0, 1), @(0, 0))
        for ($k = 0; $k -lt $pts.Count; $k++) {
            $geom += '<Row T="RelLineTo" IX="' + ($k + 2) + '">' + (VCell 'X' $pts[$k][0]) + (VCell 'Y' $pts[$k][1]) + '</Row>'
        }
        $geom += '</Section>'
    }
    $char = '<Section N="Character"><Row IX="0">' + (VCell 'Size' ($size / 72.0)) +
            (VCell 'Style' ([int]$bold)) + (VCell 'Color' '#000000') + '</Row></Section>'
    $para = '<Section N="Paragraph"><Row IX="0">' + (VCell 'HorzAlign' 1) + '</Row></Section>'
    $xml = '<Shape ID="' + $sid + '" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">' +
           $cells + $char + $para + $geom + '<Text>' + (Esc $text) + '</Text></Shape>'
    $svg = @{ kind = 'box'; cx = $cx; cy = $cy; w = $w; h = $h; text = $text; fill = $fill; line = $line;
              size = $size; bold = $bold; ellipse = $ellipse; valign = $valign; lw = $lw; r = $rounding }
    if ($back) { $pg.back.Add($xml); $pg.svgBack.Add($svg) } else { $pg.front.Add($xml); $pg.svgFront.Add($svg) }
    return @{ id = $sid; cx = $cx; cy = $cy; w = $w; h = $h }
}

function Get-Edge([double]$cx, [double]$cy, [double]$w, [double]$h, [double]$tx, [double]$ty) {
    # point de sortie sur le bord de la forme, en direction de (tx, ty)
    $dx = $tx - $cx; $dy = $ty - $cy
    if ($dx -eq 0 -and $dy -eq 0) { return ,@($cx, $cy) }
    $sx = [double]::PositiveInfinity; $sy = [double]::PositiveInfinity
    if ($dx -ne 0) { $sx = ($w / 2) / [math]::Abs($dx) }
    if ($dy -ne 0) { $sy = ($h / 2) / [math]::Abs($dy) }
    $s = [math]::Min($sx, $sy)
    return ,@(($cx + $dx * $s), ($cy + $dy * $s))
}

function Add-Connector($pg, $a, $b, [string]$text = '', [string]$color = '#404040', [double]$lw = 1.5,
                       [bool]$dash = $false, [bool]$arrow = $false) {
    # Connecteur colle (glue dynamique) entre deux formes
    $p1 = Get-Edge $a.cx $a.cy $a.w $a.h $b.cx $b.cy
    $p2 = Get-Edge $b.cx $b.cy $b.w $b.h $a.cx $a.cy
    $bx0 = [double]$p1[0]; $by0 = [double]$p1[1]; $ex0 = [double]$p2[0]; $ey0 = [double]$p2[1]
    $sid = $pg.nextId
    $pg.nextId = $sid + 1
    $w = $ex0 - $bx0; $h = $ey0 - $by0
    $lp = 1; if ($dash) { $lp = 2 }
    $ea = 0; if ($arrow) { $ea = 4 }
    $cells = @(
        (VCell 'PinX' (($bx0 + $ex0) / 2) 'GUARD((BeginX+EndX)/2)'),
        (VCell 'PinY' (($by0 + $ey0) / 2) 'GUARD((BeginY+EndY)/2)'),
        (VCell 'Width' $w 'GUARD(EndX-BeginX)'), (VCell 'Height' $h 'GUARD(EndY-BeginY)'),
        (VCell 'LocPinX' ($w / 2) 'GUARD(Width*0.5)'), (VCell 'LocPinY' ($h / 2) 'GUARD(Height*0.5)'),
        (VCell 'BeginX' $bx0 '_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)'),
        (VCell 'BeginY' $by0 '_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)'),
        (VCell 'EndX' $ex0 '_WALKGLUE(EndTrigger,BegTrigger,WalkPreference)'),
        (VCell 'EndY' $ey0 '_WALKGLUE(EndTrigger,BegTrigger,WalkPreference)'),
        (VCell 'BegTrigger' 2 ('_XFTRIGGER(Sheet.' + $a.id + '!EventXFMod)')),
        (VCell 'EndTrigger' 2 ('_XFTRIGGER(Sheet.' + $b.id + '!EventXFMod)')),
        (VCell 'ObjType' 2), (VCell 'ShapeRouteStyle' 16), (VCell 'ConFixedCode' 6),
        (VCell 'LineColor' $color), (VCell 'LineWeight' ($lw / 72.0)),
        (VCell 'LinePattern' $lp), (VCell 'BeginArrow' 0), (VCell 'EndArrow' $ea), (VCell 'EndArrowSize' 2),
        (VCell 'FillPattern' 0), (VCell 'TextBkgnd' '#FFFFFF'),
        (VCell 'TxtPinX' ($w / 2) 'Width*0.5'), (VCell 'TxtPinY' ($h / 2) 'Height*0.5'),
        (VCell 'TxtWidth' 1.4), (VCell 'TxtHeight' 0.4),
        (VCell 'TxtLocPinX' 0.7 'TxtWidth*0.5'), (VCell 'TxtLocPinY' 0.2 'TxtHeight*0.5')
    ) -join ''
    $geom = '<Section N="Geometry" IX="0">' + (VCell 'NoFill' 1) + (VCell 'NoLine' 0) + (VCell 'NoShow' 0) +
            '<Row T="MoveTo" IX="1">' + (VCell 'X' 0) + (VCell 'Y' 0) + '</Row>' +
            '<Row T="LineTo" IX="2">' + (VCell 'X' $w 'Width') + (VCell 'Y' $h 'Height') + '</Row></Section>'
    $char = '<Section N="Character"><Row IX="0">' + (VCell 'Size' (7.5 / 72.0)) + (VCell 'Color' '#202020') + '</Row></Section>'
    $pg.back.Add('<Shape ID="' + $sid + '" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">' +
                 $cells + $char + $geom + '<Text>' + (Esc $text) + '</Text></Shape>')
    $pg.svgBack.Add(@{ kind = 'line'; x1 = $bx0; y1 = $by0; x2 = $ex0; y2 = $ey0; text = $text; color = $color;
                       lw = $lw; dash = $dash; arrow = $arrow })
    $pg.connects.Add('<Connect FromSheet="' + $sid + '" FromCell="BeginX" FromPart="9" ToSheet="' + $a.id + '" ToCell="PinX" ToPart="3"/>')
    $pg.connects.Add('<Connect FromSheet="' + $sid + '" FromCell="EndX" FromPart="12" ToSheet="' + $b.id + '" ToCell="PinX" ToPart="3"/>')
    return $sid
}

function VPage-Xml($pg) {
    return '<?xml version="1.0" encoding="utf-8"?><PageContents xmlns="' + $NS + '" xmlns:r="' + $NSR +
           '" xml:space="preserve"><Shapes>' + ($pg.back -join '') + ($pg.front -join '') + '</Shapes><Connects>' +
           ($pg.connects -join '') + '</Connects></PageContents>'
}

function Write-Zip([string]$path, $entries) {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive -ArgumentList $fs, ([System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in @($entries.Keys)) {
                $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $stream = $e.Open()
                $sw = New-Object System.IO.StreamWriter -ArgumentList $stream, $Utf8NoBom
                try { $sw.Write([string]$entries[$name]) } finally { $sw.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Write-Vsdx($pages, [string]$path, [string]$title = 'Topologie Check Point') {
    $ct = '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>' +
          '<Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>' +
          '<Override PartName="/visio/windows.xml" ContentType="application/vnd.ms-visio.windows+xml"/>'
    for ($i = 1; $i -le $pages.Count; $i++) {
        $ct += '<Override PartName="/visio/pages/page' + $i + '.xml" ContentType="application/vnd.ms-visio.page+xml"/>'
    }
    $ct += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
           '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
           '</Types>'
    $rels = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="' + $RELV + 'document" Target="visio/document.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
            '</Relationships>'
    $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $Inv)
    $core = '<?xml version="1.0" encoding="utf-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
            '<dc:title>' + (Esc $title) + '</dc:title><dc:creator>cp_topo.ps1</dc:creator>' +
            '<dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created></cp:coreProperties>'
    $app = '<?xml version="1.0" encoding="utf-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
           '<Application>Microsoft Visio</Application><AppVersion>16.0000</AppVersion></Properties>'
    $style0 = (@(
            (VCell 'EnableLineProps' 1), (VCell 'EnableFillProps' 1), (VCell 'EnableTextProps' 1), (VCell 'HideForApply' 0),
            (VCell 'LineWeight' 0.01041666666666667), (VCell 'LineColor' 0), (VCell 'LinePattern' 1), (VCell 'Rounding' 0),
            (VCell 'EndArrowSize' 2), (VCell 'BeginArrow' 0), (VCell 'EndArrow' 0), (VCell 'LineCap' 0),
            (VCell 'BeginArrowSize' 2), (VCell 'LineColorTrans' 0), (VCell 'FillForegnd' 1), (VCell 'FillBkgnd' 0),
            (VCell 'FillPattern' 1), (VCell 'ShdwForegnd' 0), (VCell 'ShdwPattern' 0), (VCell 'FillForegndTrans' 0),
            (VCell 'FillBkgndTrans' 0), (VCell 'ShdwForegndTrans' 0), (VCell 'LeftMargin' 0), (VCell 'RightMargin' 0),
            (VCell 'TopMargin' 0), (VCell 'BottomMargin' 0), (VCell 'VerticalAlign' 1), (VCell 'TextBkgnd' 0),
            (VCell 'TextDirection' 0), (VCell 'TextBkgndTrans' 0)) -join '') +
        '<Section N="Character"><Row IX="0">' + (VCell 'Color' 0) + (VCell 'Style' 0) + (VCell 'Size' 0.1666666666666667) +
        (VCell 'Case' 0) + '</Row></Section>' +
        '<Section N="Paragraph"><Row IX="0">' + (VCell 'HorzAlign' 1) + '</Row></Section>'
    $doc = '<?xml version="1.0" encoding="utf-8"?><VisioDocument xmlns="' + $NS + '" xmlns:r="' + $NSR + '" xml:space="preserve">' +
           '<DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="0">' +
           '<GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions>' +
           '<DynamicGridEnabled>0</DynamicGridEnabled><ProtectStyles>0</ProtectStyles><ProtectShapes>0</ProtectShapes>' +
           '<ProtectMasters>0</ProtectMasters><ProtectBkgnds>0</ProtectBkgnds></DocumentSettings>' +
           '<Colors><ColorEntry IX="0" RGB="#000000"/><ColorEntry IX="1" RGB="#FFFFFF"/></Colors>' +
           '<StyleSheets><StyleSheet ID="0" NameU="No Style" IsCustomNameU="1" Name="No Style" IsCustomName="1">' + $style0 +
           '</StyleSheet></StyleSheets>' +
           '<DocumentSheet NameU="TheDoc" IsCustomNameU="1" Name="TheDoc" IsCustomName="1" LineStyle="0" FillStyle="0" TextStyle="0">' +
           (VCell 'OutputFormat' 0) + (VCell 'LockPreview' 0) + (VCell 'AddMarkup' 0) + (VCell 'ViewMarkup' 0) +
           '</DocumentSheet></VisioDocument>'
    $p0 = $pages[0]
    $windows = '<?xml version="1.0" encoding="utf-8"?><Windows ClientWidth="1600" ClientHeight="900" xmlns="' + $NS +
               '" xmlns:r="' + $NSR + '" xml:space="preserve"><Window ID="0" WindowType="Drawing" WindowState="1073741824" ' +
               'WindowLeft="-8" WindowTop="-31" WindowWidth="1616" WindowHeight="939" ContainerType="Page" Page="0" ' +
               'ViewScale="-1" ViewCenterX="' + (Fmt ($p0.width / 2)) + '" ViewCenterY="' + (Fmt ($p0.height / 2)) + '">' +
               '<ShowRulers>1</ShowRulers><ShowGrid>1</ShowGrid><ShowPageBreaks>0</ShowPageBreaks><ShowGuides>1</ShowGuides>' +
               '<ShowConnectionPoints>1</ShowConnectionPoints><GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings>' +
               '<SnapExtensions>34</SnapExtensions><SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled>' +
               '<TabSplitterPos>0.5</TabSplitterPos></Window></Windows>'
    $docRels = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
               '<Relationship Id="rId1" Type="' + $RELV + 'pages" Target="pages/pages.xml"/>' +
               '<Relationship Id="rId2" Type="' + $RELV + 'windows" Target="windows.xml"/></Relationships>'
    $pagesXml = '<?xml version="1.0" encoding="utf-8"?><Pages xmlns="' + $NS + '" xmlns:r="' + $NSR + '" xml:space="preserve">'
    $pagesRels = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $p = $pages[$i]
        $pagesXml += '<Page ID="' + $i + '" NameU="' + (Esc $p.name) + '" Name="' + (Esc $p.name) + '" ViewScale="-1" ViewCenterX="' +
                     (Fmt ($p.width / 2)) + '" ViewCenterY="' + (Fmt ($p.height / 2)) + '">' +
                     '<PageSheet LineStyle="0" FillStyle="0" TextStyle="0">' +
                     (VCell 'PageWidth' $p.width '' 'IN') + (VCell 'PageHeight' $p.height '' 'IN') +
                     (VCell 'PageScale' 1 '' 'IN_F') + (VCell 'DrawingScale' 1 '' 'IN_F') +
                     (VCell 'DrawingSizeType' 3) + (VCell 'DrawingScaleType' 0) + (VCell 'PrintPageOrientation' 2) +
                     '</PageSheet><Rel r:id="rId' + ($i + 1) + '"/></Page>'
        $pagesRels += '<Relationship Id="rId' + ($i + 1) + '" Type="' + $RELV + 'page" Target="page' + ($i + 1) + '.xml"/>'
    }
    $pagesXml += '</Pages>'
    $pagesRels += '</Relationships>'

    $entries = [ordered]@{}
    $entries['[Content_Types].xml'] = $ct
    $entries['_rels/.rels'] = $rels
    $entries['docProps/core.xml'] = $core
    $entries['docProps/app.xml'] = $app
    $entries['visio/document.xml'] = $doc
    $entries['visio/_rels/document.xml.rels'] = $docRels
    $entries['visio/windows.xml'] = $windows
    $entries['visio/pages/pages.xml'] = $pagesXml
    $entries['visio/pages/_rels/pages.xml.rels'] = $pagesRels
    for ($i = 0; $i -lt $pages.Count; $i++) { $entries['visio/pages/page' + ($i + 1) + '.xml'] = (VPage-Xml $pages[$i]) }
    Write-Zip $path $entries
}

# ---------------------------------------------------------------------------
# 6b. Rendu SVG de secours (Visio : Fichier > Ouvrir > *.svg)
# ---------------------------------------------------------------------------

function Svg-Text([double]$cx, [double]$cy, [string[]]$lines, [double]$size, [bool]$bold, $top) {
    $lh = $size * 1.25
    if ($null -ne $top) { $y0 = [double]$top + $lh } else { $y0 = $cy - $lh * ($lines.Count - 1) / 2 + $size * 0.35 }
    $b = ''; if ($bold) { $b = ' font-weight="bold"' }
    $s = '<text x="' + (F1 $cx) + '" y="' + (F1 $y0) + '" font-size="' + (F1 $size) + '" text-anchor="middle"' + $b + '>'
    for ($k = 0; $k -lt $lines.Count; $k++) {
        $dy = '0'; if ($k -gt 0) { $dy = F1 $lh }
        $s += '<tspan x="' + (F1 $cx) + '" dy="' + $dy + '">' + (Esc $lines[$k]) + '</tspan>'
    }
    return $s + '</text>'
}

function Write-Svg($pg, [string]$path) {
    $D = 96.0
    $PW = $pg.width * $D; $PH = $pg.height * $D
    $o = New-Object System.Collections.Generic.List[string]
    $o.Add('<?xml version="1.0" encoding="UTF-8"?>')
    $o.Add('<svg xmlns="http://www.w3.org/2000/svg" width="' + (F1 $PW) + '" height="' + (F1 $PH) + '" viewBox="0 0 ' +
           (F1 $PW) + ' ' + (F1 $PH) + '" font-family="Calibri, Arial, sans-serif">')
    $o.Add('<defs><marker id="arr" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="7" markerHeight="7" ' +
           'orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="context-stroke"/></marker></defs>')
    $o.Add('<rect width="100%" height="100%" fill="#FFFFFF"/>')
    foreach ($it in (@($pg.svgBack) + @($pg.svgFront))) {
        if ($it.kind -eq 'box') {
            $cx = $it.cx * $D; $cy = $PH - $it.cy * $D; $bw = $it.w * $D; $bh = $it.h * $D
            if ($it.ellipse) {
                $o.Add('<ellipse cx="' + (F1 $cx) + '" cy="' + (F1 $cy) + '" rx="' + (F1 ($bw / 2)) + '" ry="' + (F1 ($bh / 2)) +
                       '" fill="' + $it.fill + '" stroke="' + $it.line + '" stroke-width="' + (F1 ($it.lw * 1.33)) + '"/>')
            } else {
                $o.Add('<rect x="' + (F1 ($cx - $bw / 2)) + '" y="' + (F1 ($cy - $bh / 2)) + '" width="' + (F1 $bw) + '" height="' +
                       (F1 $bh) + '" rx="' + (F1 ($it.r * $D)) + '" fill="' + $it.fill + '" stroke="' + $it.line +
                       '" stroke-width="' + (F1 ($it.lw * 1.33)) + '"/>')
            }
            if ($it.text) {
                $top = $null; if ($it.valign -eq 0) { $top = $cy - $bh / 2 + 2 }
                $o.Add((Svg-Text $cx $cy ($it.text -split "`n") ($it.size * 1.33) $it.bold $top))
            }
        } else {
            $x1 = $it.x1 * $D; $y1 = $PH - $it.y1 * $D; $x2 = $it.x2 * $D; $y2 = $PH - $it.y2 * $D
            $extra = ''
            if ($it.dash) { $extra += ' stroke-dasharray="6,4"' }
            if ($it.arrow) { $extra += ' marker-end="url(#arr)"' }
            $o.Add('<line x1="' + (F1 $x1) + '" y1="' + (F1 $y1) + '" x2="' + (F1 $x2) + '" y2="' + (F1 $y2) + '" stroke="' +
                   $it.color + '" stroke-width="' + (F1 ($it.lw * 1.33)) + '"' + $extra + '/>')
            if ($it.text) {
                $lines = $it.text -split "`n"
                $mx = ($x1 + $x2) / 2; $my = ($y1 + $y2) / 2
                $maxLen = 0; foreach ($l in $lines) { if ($l.Length -gt $maxLen) { $maxLen = $l.Length } }
                $tw = $maxLen * 5.6 + 8; $th = $lines.Count * 12.5 + 4
                $o.Add('<rect x="' + (F1 ($mx - $tw / 2)) + '" y="' + (F1 ($my - $th / 2)) + '" width="' + (F1 $tw) +
                       '" height="' + (F1 $th) + '" fill="#FFFFFF"/>')
                $o.Add((Svg-Text $mx $my $lines 10 $false $null))
            }
        }
    }
    $o.Add('</svg>')
    Write-Text $path ($o -join "`n") $Utf8NoBom
}

# ---------------------------------------------------------------------------
# 7. Mise en page de la topologie
# ---------------------------------------------------------------------------

function RowW([int]$k, [double]$w, [double]$gap) { return $k * $w + [math]::Max($k - 1, 0) * $gap }

function Mem-Lines($m) {
    $ls = New-Object System.Collections.Generic.List[string]
    $ls.Add($m.name)
    if ($m.ip) { $ls.Add("Principale : $($m.ip)") }
    $k = 0
    foreach ($i in $m.interfaces) {
        if ($k -ge 12) { break }
        $p = '?'; if ($null -ne $i.prefix) { $p = $i.prefix }
        $ls.Add("$($i.name)  $($i.ip)/$p")
        $k++
    }
    if ($m.interfaces.Count -gt 12) { $ls.Add("... +$($m.interfaces.Count - 12)") }
    return ,$ls
}

function Net-Text($n) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add($n.network)
    $lines.Add((Label-Of $n.role))
    $k = 0
    foreach ($v in $n.vip) { if ($k -ge 2) { break }; $lines.Add(('VIP {0} ({1})' -f $v.ip, $v.iface)); $k++ }
    return ($lines -join "`n")
}

function Conn-Label($n) {
    $ifs = @((@($n.vip | ForEach-Object { $_.iface }) + @($n.members | ForEach-Object { $_.iface })) | Sort-Object -Unique)
    $mem = (@($n.members | ForEach-Object { Short-Ip $_.ip }) -join ' ')
    if ($mem) { return ($ifs -join '/') + "`n" + $mem }
    return ($ifs -join '/')
}

function Draw-Topology($topo) {
    $cl = $topo.cluster; $members = $topo.members; $nets = $topo.networks
    $pname = [string]$cl.name; if ($pname.Length -gt 30) { $pname = $pname.Substring(0, 30) }
    $pg = New-VPage $pname
    $NET_W = 2.3; $NET_H = 0.85; $GAP = 0.55; $MEM_W = 2.3; $margin = 0.8

    $top = @($nets | Where-Object { $_.role -eq 'external' })
    $side = @($nets | Where-Object { ($_.role -eq 'sync' -or $_.role -eq 'mgmt') -and $_.vip.Count -eq 0 })
    $bottom = @($nets | Where-Object { $top -notcontains $_ -and $side -notcontains $_ })

    $memH = 1.0
    foreach ($m in $members) { $memH = [math]::Max($memH, 0.4 + 0.15 * (Mem-Lines $m).Count) }
    $nMem = [math]::Max($members.Count, 1)
    $clW = [math]::Max($nMem * ($MEM_W + 0.4) + 0.4, 4.0)
    $clH = $memH + 1.0
    if ($side.Count -gt 0) { $clH += $NET_H + 0.9 }
    $clW = [math]::Max($clW, $side.Count * ($NET_W + 0.4) + 0.4)
    $contentW = [math]::Max([math]::Max((RowW $top.Count $NET_W $GAP), (RowW $bottom.Count $NET_W $GAP)), $clW)
    $hasRtTop = @($top | Where-Object { $_.routers.Count -gt 0 }).Count -gt 0
    $hasRtBot = @($bottom | Where-Object { $_.routers.Count -gt 0 }).Count -gt 0

    $pg.width = $contentW + 2 * $margin
    $y = $margin
    $yRtBot = 0.0
    if ($hasRtBot) { $yRtBot = $y + 0.5; $y += 1.6 }
    $yBot = $y + $NET_H / 2
    $y += $NET_H + 1.6
    $yCl = $y + $clH / 2
    $y += $clH + 1.6
    $yTop = $y + $NET_H / 2
    $y += $NET_H
    $yRtTop = 0.0
    if ($hasRtTop) { $yRtTop = $y + 1.1; $y += 1.6 }
    $y += 1.3   # titre + legende
    $pg.height = $y + $margin * 0.5
    $cxMid = $margin + $contentW / 2

    # Titre
    $title = 'Topologie du cluster {0} {1} {2} membre(s), {3} réseau(x)' -f $cl.name, [char]0x2014, $members.Count, $nets.Count
    $null = Add-Box $pg ($pg.width / 2) ($pg.height - 0.55) ([math]::Min($pg.width - 1, 9)) 0.5 $title '#FFFFFF' '#FFFFFF' -size 14 -bold $true -lw 0.1

    # Cluster (cadre) + membres
    $mode = ''
    foreach ($k in @($cl.props.Keys)) { if ($k.Contains('mode') -or $k.Contains('high availability')) { $mode = $cl.props[$k]; break } }
    $clText = [string]$cl.name
    if ($cl.type) { $clText += " ($($cl.type))" }
    if ($mode) { $clText += "`n$mode" }
    $sc = Get-Style 'cluster'
    $clShape = Add-Box $pg $cxMid $yCl $clW $clH $clText $sc[0] $sc[1] -size 11 -bold $true -valign 0 -lw 2 -back $true
    $memShapes = @{}
    $x0 = $cxMid - ($nMem * ($MEM_W + 0.4) - 0.4) / 2 + $MEM_W / 2
    $memY = $yCl + $clH / 2 - 0.75 - $memH / 2
    $sm = Get-Style 'member'
    for ($k = 0; $k -lt $members.Count; $k++) {
        $m = $members[$k]
        $memShapes[$m.name] = Add-Box $pg ($x0 + $k * ($MEM_W + 0.4)) $memY $MEM_W $memH ((Mem-Lines $m) -join "`n") $sm[0] $sm[1] -size 7.5 -lw 1.25
    }

    # Reseaux externes (haut) et internes/DMZ (bas)
    $placed = New-Object System.Collections.Generic.List[object]
    foreach ($pair in @(@{ row = $top; y = $yTop }, @{ row = $bottom; y = $yBot })) {
        $xx = $cxMid - (RowW $pair.row.Count $NET_W $GAP) / 2 + $NET_W / 2
        foreach ($n in $pair.row) {
            $stl = Get-Style $n.role
            $s = Add-Box $pg $xx $pair.y $NET_W $NET_H (Net-Text $n) $stl[0] $stl[1] -size 8.5
            $placed.Add(@{ n = $n; s = $s; y = $pair.y })
            $xx += $NET_W + $GAP
        }
    }
    $sr = Get-Style 'router'
    foreach ($pl in $placed) {
        $n = $pl.n; $s = $pl.s; $stl = Get-Style $n.role
        $target = $clShape
        if ($n.vip.Count -eq 0 -and $n.members.Count -gt 0 -and $memShapes.ContainsKey($n.members[0].member)) {
            $target = $memShapes[$n.members[0].member]
        }
        $null = Add-Connector $pg $s $target (Conn-Label $n) -color $stl[1] -lw 2
        if ($n.routers.Count -gt 0) {
            if ($pl.y -eq $yTop) { $ry = $yRtTop } else { $ry = $yRtBot }
            $kk = $n.routers.Count; $j = 0
            foreach ($nh in @($n.routers.Keys)) {
                $dests = $n.routers[$nh]
                $rx = $s.cx + ($j - ($kk - 1) / 2) * 1.5
                $dl = (@($dests | Select-Object -First 5) -join "`n")
                if ($dests.Count -gt 5) { $dl += "`n..." }
                $r = Add-Box $pg $rx $ry 1.4 0.75 ("Routeur $nh`n$dl") $sr[0] $sr[1] -size 7 -ellipse $true
                $null = Add-Connector $pg $s $r '' -color $sr[1] -lw 1
                $j++
            }
        }
    }

    # Reseaux synchro / gestion dans le cadre du cluster, sous les membres
    $xs = $cxMid - ($side.Count * ($NET_W + 0.4) - 0.4) / 2 + $NET_W / 2
    $yy = $yCl - $clH / 2 + 0.25 + $NET_H / 2
    foreach ($n in $side) {
        $stl = Get-Style $n.role
        $s = Add-Box $pg $xs $yy $NET_W $NET_H (Net-Text $n) $stl[0] $stl[1] -size 8.5
        foreach ($mm in $n.members) {
            if ($memShapes.ContainsKey($mm.member)) {
                $null = Add-Connector $pg $s $memShapes[$mm.member] ('{0} {1}' -f $mm.iface, (Short-Ip $mm.ip)) -color $stl[1] -lw 1.25 -dash $true
            }
        }
        $xs += $NET_W + 0.4
    }

    # Legende (en haut a gauche)
    $k = 0
    foreach ($cls in @('external', 'internal', 'dmz', 'sync', 'mgmt')) {
        $stl = Get-Style $cls
        $null = Add-Box $pg ($margin * 0.5 + 0.45 + $k * 1.0) ($pg.height - 1.15) 0.9 0.26 (Label-Of $cls) $stl[0] $stl[1] -size 7
        $k++
    }
    return $pg
}

# ---------------------------------------------------------------------------
# 7b. Analyse des couches (layers) et de leurs regles
# ---------------------------------------------------------------------------

function Is-Any([string]$v) {
    if (-not $v) { return $false }
    $vt = $v.Trim()
    if ($vt.Contains("`n")) { return $false }
    return [regex]::IsMatch($vt, '^(any|\*any|tout)$', 'IgnoreCase')
}

function Action-Kind([string]$a) {
    $a = Norm $a
    if ($a -match 'drop|reject|block|refus|bloqu') { return 'drop' }
    if ($a -match 'accept|allow|autoris') { return 'accept' }
    if ($a -match 'inline|apply layer|layer') { return 'inline' }
    return 'other'
}

function Analyse-Layers($model) {
    $out = New-Object System.Collections.Generic.List[object]
    $names = @($model.layers.Keys)
    foreach ($ln in $names) {
        $rules = @($model.rules | Where-Object { $_['layer'] -eq $ln })
        $st = [ordered]@{ name = $ln; rules = $rules.Count; accept = 0; drop = 0; inline = 0; other = 0; disabled = 0
            sections = (New-Object System.Collections.Generic.List[string])
            children = (New-Object System.Collections.Generic.List[object])
            checks = (New-Object System.Collections.Generic.List[string]); inline_layer = $false }
        foreach ($r in $rules) {
            $act = [string]$r['action']
            $k = Action-Kind $act
            $child = $null
            foreach ($n in $names) {
                if ($n -ne $ln -and [regex]::IsMatch((Norm $act), '(?<!\w)' + [regex]::Escape((Norm $n)) + '(?!\w)')) { $child = $n; break }
            }
            $no = '?'; if ($r.Contains('no') -and $r['no']) { $no = $r['no'] }
            if ($child) { $k = 'inline'; $st.children.Add(@{ layer = $child; rule = $no }) }
            $st[$k] = $st[$k] + 1
            if ($r['disabled']) { $st['disabled'] = $st['disabled'] + 1 }
            if ($r['section'] -and -not $st.sections.Contains($r['section'])) { $st.sections.Add($r['section']) }
            if ($r['disabled']) { continue }
            $tag = "règle $no"; if ($r['rname']) { $tag += " ($($r['rname']))" }
            $arrow = [char]0x2192
            if ($k -eq 'accept' -and (Is-Any $r['src']) -and (Is-Any $r['dst']) -and ((Is-Any $r['svc']) -or -not $r['svc'])) {
                $st.checks.Add("$tag : Any $arrow Any $arrow Any en Accept")
            } elseif ($k -eq 'accept' -and ((Is-Any $r['src']) -or (Is-Any $r['dst'])) -and (Is-Any $r['svc'])) {
                $st.checks.Add("$tag : Accept avec source ou destination Any et service Any")
            }
            if ($r.Contains('track') -and ($k -eq 'accept' -or $k -eq 'drop') -and @('', 'none', 'aucun', '-') -contains (Norm $r['track'])) {
                $st.checks.Add("$tag : aucune journalisation")
            }
            if ($r.Contains('hits') -and [regex]::IsMatch([string]$r['hits'], '^\s*0\s*(\(.*\))?\s*$')) {
                $st.checks.Add("$tag : 0 hit (règle inutilisée ?)")
            }
        }
        if ($st['disabled'] -gt 0) { $st.checks.Add(('{0} règle(s) désactivée(s)' -f $st['disabled'])) }
        $enabled = @($rules | Where-Object { -not $_['disabled'] })
        if ($enabled.Count -gt 0) {
            $last = $enabled[$enabled.Count - 1]
            if ((Action-Kind $last['action']) -ne 'drop' -or -not ((Is-Any $last['src']) -and (Is-Any $last['dst']))) {
                $lno = '?'; if ($last['no']) { $lno = $last['no'] }
                $st.checks.Add(('Pas de règle de nettoyage (Any/Any/Drop) en fin de couche (dernière : règle {0}, {1})' -f $lno, $last['action']))
            }
        }
        $out.Add($st)
    }
    $childNames = @(foreach ($s in $out) { foreach ($c in $s.children) { $c.layer } })
    foreach ($s in $out) { $s['inline_layer'] = ($childNames -contains $s.name) }
    return ,$out
}

function Layer-Text($s) {
    $t = ('{0}' -f $s.name) + "`n" + ('{0} règles : {1} accept / {2} drop' -f $s.rules, $s.accept, $s.drop)
    if ($s['inline'] -gt 0 -or $s.children.Count -gt 0) { $t += (' / {0} inline' -f [math]::Max($s['inline'], $s.children.Count)) }
    if ($s['disabled'] -gt 0) { $t += "`n" + ('{0} désactivée(s)' -f $s['disabled']) }
    if ($s.sections.Count -gt 0) { $t += "`n" + ('{0} section(s)' -f $s.sections.Count) }
    if ($s.checks.Count -gt 0) { $t += "`n" + [char]0x26A0 + (" {0} point(s) d'attention" -f $s.checks.Count) }
    return $t
}

function Draw-Layers($stats, [string]$title) {
    $pg = New-VPage 'Couches'
    $BW = 3.0; $BH = 1.15; $VG = 0.8
    $ordered = @($stats | Where-Object { -not $_['inline_layer'] })
    $inl = @($stats | Where-Object { $_['inline_layer'] })
    $rowsN = [math]::Max($ordered.Count, 1)
    $pg.width = 2 * $BW + 4.5
    $pg.height = $rowsN * ($BH + $VG) + 2.4 + [math]::Max(0, $inl.Count - $rowsN) * ($BH + 0.4)
    $null = Add-Box $pg ($pg.width / 2) ($pg.height - 0.5) ($pg.width - 1) 0.5 ('Couches de la politique ' + [char]0x2014 + ' ' + $title) '#FFFFFF' '#FFFFFF' -size 14 -bold $true -lw 0.1
    $shapes = @{}
    $x0 = 0.9 + $BW / 2
    $yy = $pg.height - 1.4 - $BH / 2
    $prev = $null; $k = 0
    foreach ($s in $ordered) {
        $k++
        $sh = Add-Box $pg $x0 $yy $BW $BH ("$k. " + (Layer-Text $s)) '#DAE8FC' '#6C8EBF' -size 9
        $shapes[$s.name] = $sh
        if ($null -ne $prev) { $null = Add-Connector $pg $prev $sh 'Accept' -color '#6C8EBF' -lw 1.5 -arrow $true }
        $prev = $sh
        $yy -= ($BH + $VG)
    }
    $xi = $x0 + $BW + 1.8
    $yi = $pg.height - 1.4 - $BH / 2
    foreach ($s in $inl) {
        $sh = Add-Box $pg $xi $yi $BW $BH ('Inline : ' + (Layer-Text $s)) '#FFF2CC' '#D6B656' -size 9
        $shapes[$s.name] = $sh
        $yi -= ($BH + 0.4)
    }
    foreach ($s in $stats) {
        foreach ($c in $s.children) {
            if ($shapes.ContainsKey($s.name) -and $shapes.ContainsKey($c.layer)) {
                $null = Add-Connector $pg $shapes[$s.name] $shapes[$c.layer] ('règle ' + $c.rule) -color '#D6B656' -lw 1.25 -dash $true -arrow $true
            }
        }
    }
    return $pg
}

function Write-RulesCsv($model, [string]$path) {
    $cols = @('layer', 'section', 'no', 'rname', 'src', 'dst', 'vpn', 'svc', 'content', 'action', 'track', 'install', 'time', 'hits', 'comment', 'disabled')
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Couche;Section;N°;Nom;Source;Destination;VPN;Services;Contenu;Action;Track;Installé sur;Horaire;Hits;Commentaire;Désactivée')
    foreach ($r in $model.rules) {
        $vals = foreach ($c in $cols) {
            $v = ''; if ($r.Contains($c)) { $v = [string]$r[$c] }
            '"' + ($v -replace "`n", ', ').Replace('"', '""') + '"'
        }
        $lines.Add(($vals -join ';'))
    }
    Write-Text $path (($lines -join "`r`n") + "`r`n") $Utf8Bom
}

# ---------------------------------------------------------------------------
# 8. Rapport Markdown
# ---------------------------------------------------------------------------

function Report-Md($topos, $model, $stats) {
    $L = New-Object System.Collections.Generic.List[string]
    $L.Add('# Analyse de configuration Check Point'); $L.Add('')
    $arrow = [char]0x2192
    if ($stats.Count -gt 0) {
        $L.Add('## Couches de politique'); $L.Add('')
        $L.Add('| # | Couche | Type | Règles | Accept | Drop | Inline | Désactivées | Sections |')
        $L.Add('|---|---|---|---|---|---|---|---|---|')
        $k = 0
        foreach ($s in $stats) {
            $k++
            $typ = 'ordonnée'; if ($s['inline_layer']) { $typ = 'inline' }
            $inl = $s['inline']; if ($s.children.Count -gt 0) { $inl = $s.children.Count }
            $L.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} |' -f $k, $s.name, $typ, $s.rules, $s.accept, $s.drop, $inl, $s['disabled'], $s.sections.Count))
        }
        foreach ($s in $stats) {
            $L.Add(''); $L.Add("### Couche $($s.name)"); $L.Add('')
            if ($s.sections.Count -gt 0) { $L.Add('Sections : ' + ($s.sections -join ', ')); $L.Add('') }
            if ($s.checks.Count -gt 0) { foreach ($c in $s.checks) { $L.Add("- $c") } } else { $L.Add("- Aucun point d'attention") }
        }
        $L.Add('')
    }
    foreach ($t in $topos) {
        $cl = $t.cluster
        $L.Add("## Cluster $($cl.name)"); $L.Add('')
        if ($cl.type) { $L.Add("- Type : $($cl.type)") }
        if ($cl.ip) { $L.Add("- Adresse principale : $($cl.ip)") }
        $mn = (@($t.members | ForEach-Object { $_.name }) -join ', '); if (-not $mn) { $mn = 'aucun détecté' }
        $L.Add("- Membres : $mn")
        $L.Add(''); $L.Add('### Interfaces virtuelles (VIP)'); $L.Add('')
        $L.Add('| Interface | VIP | Réseau | Topologie |'); $L.Add('|---|---|---|---|')
        foreach ($i in $cl.interfaces) {
            $net = Net-Of $i; if (-not $net) { $net = '?' }
            $topo = $i.topology_raw; if (-not $topo) { $topo = Label-Of $i.topology }
            $L.Add(('| {0} | {1} | {2} | {3} |' -f $i.name, $i.ip, $net, $topo))
        }
        foreach ($m in $t.members) {
            $L.Add(''); $L.Add("### Membre $($m.name)"); $L.Add('')
            $L.Add('| Interface | IP | Réseau |'); $L.Add('|---|---|---|')
            foreach ($i in $m.interfaces) {
                $net = Net-Of $i; if (-not $net) { $net = '?' }
                $L.Add(('| {0} | {1} | {2} |' -f $i.name, $i.ip, $net))
            }
        }
        $L.Add(''); $L.Add('### Réseaux'); $L.Add('')
        $L.Add('| Réseau | Rôle | VIP | Membres | Routeurs |'); $L.Add('|---|---|---|---|---|')
        foreach ($n in $t.networks) {
            $vips = (@($n.vip | ForEach-Object { $_.ip }) -join ', ')
            $mems = (@($n.members | ForEach-Object { '{0} {1}' -f $_.member, $_.ip }) -join ', ')
            $rts = (@(foreach ($rk in @($n.routers.Keys)) { "$rk $arrow " + ($n.routers[$rk] -join ', ') }) -join '; ')
            $L.Add(('| {0} | {1} | {2} | {3} | {4} |' -f $n.network, (Label-Of $n.role), $vips, $mems, $rts))
        }
        $L.Add(''); $L.Add("### Points d'attention"); $L.Add('')
        if ($t.checks.Count -gt 0) { foreach ($c in $t.checks) { $L.Add("- $c") } } else { $L.Add('- Aucun') }
        $L.Add('')
    }
    return ($L -join "`n")
}

# ---------------------------------------------------------------------------
# 9. Programme principal
# ---------------------------------------------------------------------------

$files = New-Object System.Collections.Generic.List[string]
foreach ($p in $Path) {
    $rp = Full $p
    if (Test-Path -LiteralPath $rp -PathType Container) {
        Get-ChildItem -LiteralPath $rp -Recurse -File | Where-Object { $_.Extension -match '^\.html?$' } |
            Sort-Object FullName | ForEach-Object { $files.Add($_.FullName) }
    } elseif (Test-Path -LiteralPath $rp) {
        $files.Add($rp)
    } else {
        Write-Warning "Introuvable : $p"
    }
}
if ($files.Count -eq 0) { Write-Error 'Aucun fichier HTML trouvé.'; exit 1 }

$model = Build-Model $files

if ($ShowTables) {
    Write-Host '=== Tableaux détectés ==='
    foreach ($t in $model.log) {
        Write-Host ('[{0}] {1,-22} {2,3} lignes | {3}' -f $t.source, $t.interpretation, $t.lignes, $t.contexte)
        Write-Host ('      entete={0}  colonnes={1}' -f ((@($t.entete) | Select-Object -First 8) -join ' | '), $t.colonnes)
    }
    Write-Host '=== Objets ==='
    foreach ($o in $model.objects.Values) {
        Write-Host ('  {0,-25} kind={1,-8} type={2,-25} itf={3} membres={4}' -f $o.name, $o.kind, $o.type, $o.interfaces.Count, ($o.members -join ','))
    }
}

$allTopos = Build-Topology $model $Cluster
$topos = @($allTopos | Where-Object { $_.cluster.interfaces.Count -gt 0 -or $_.members.Count -gt 0 })
$layerStats = Analyse-Layers $model
if ($topos.Count -eq 0 -and $layerStats.Count -eq 0) {
    Write-Error 'Ni topologie ni règles trouvées. Relancez avec -ShowTables pour voir les tableaux détectés et ajustez $FieldAliases / $RuleAliases.'
    exit 1
}
if ($topos.Count -eq 0) { Write-Warning 'Aucune interface de cluster trouvée (seules les couches seront dessinées).' }

if ($Output) { $out = Full $Output }
else { $out = Join-Path (Split-Path $files[0]) ([System.IO.Path]::GetFileNameWithoutExtension($files[0]) + '_topologie.vsdx') }
$outDir = Split-Path $out
if ($outDir) { [void][System.IO.Directory]::CreateDirectory($outDir) }
$base = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($out))

$pages = New-Object System.Collections.Generic.List[object]
foreach ($t in $topos) { $pages.Add((Draw-Topology $t)) }
if ($layerStats.Count -gt 0) {
    $pages.Add((Draw-Layers $layerStats ((@($topos | ForEach-Object { $_.cluster.name })) -join ', ')))
}
Write-Vsdx $pages $out

$svgs = @()
for ($i = 0; $i -lt $pages.Count; $i++) {
    $sp = '{0}_page{1}.svg' -f $base, ($i + 1)
    Write-Svg $pages[$i] $sp
    $svgs += $sp
}
Write-Host ('SVG     : ' + ($svgs -join ', '))

if ($Report) { $rep = Full $Report } else { $rep = $base + '.md' }
Write-Text $rep (Report-Md $topos $model $layerStats) $Utf8NoBom

if ($model.rules.Count -gt 0) {
    $csv = $base + '_regles.csv'
    Write-RulesCsv $model $csv
    Write-Host ('Règles  : {0} ({1} règles, {2} couche(s))' -f $csv, $model.rules.Count, $layerStats.Count)
}
if ($Json) {
    $cl = @(foreach ($t in $topos) { [ordered]@{ cluster = $t.cluster; members = $t.members; networks = $t.networks; checks = $t.checks } })
    $data = [ordered]@{ clusters = $cl; layers = $layerStats; rules = $model.rules }
    Write-Text (Full $Json) ($data | ConvertTo-Json -Depth 10) $Utf8NoBom
}

foreach ($t in $topos) {
    Write-Host ('Cluster {0} : {1} membre(s), {2} réseau(x), {3} point(s) d''attention' -f $t.cluster.name, $t.members.Count, $t.networks.Count, $t.checks.Count)
}
Write-Host ('Visio   : ' + $out)
Write-Host ('Rapport : ' + $rep)
