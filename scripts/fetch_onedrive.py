#!/usr/bin/env python3
"""Unduh file xlsx dari link berbagi OneDrive pribadi (format lama maupun baru).
Usage: fetch_onedrive.py <share_url> <output.xlsx>. Tidak mencetak link/token."""
import base64, sys, zipfile, urllib.request, urllib.error
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

URL, OUT = "".join(sys.argv[1].split()), sys.argv[2]
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

class Keep(urllib.request.HTTPRedirectHandler):
    pass
opener = urllib.request.build_opener(Keep)

def get(url, label):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    try:
        r = opener.open(req, timeout=60)
        data, final, code = r.read(), r.geturl(), r.status
    except urllib.error.HTTPError as e:
        print(f"{label}: HTTP {e.code}"); return None, e.url if hasattr(e, "url") else url
    except Exception as e:
        print(f"{label}: error {type(e).__name__}"); return None, url
    ok = zipfile.is_zipfile(__import__("io").BytesIO(data))
    print(f"{label}: HTTP {code}, {len(data)} byte, xlsx={ok}")
    return (data if ok else None), final

def with_param(url, **kw):
    p = urlparse(url); q = parse_qs(p.query, keep_blank_values=True)
    for k, v in kw.items(): q[k] = [v]
    return urlunparse(p._replace(query=urlencode(q, doseq=True)))

def describe(url):
    p = urlparse(url); q = parse_qs(p.query)
    print(f"  tujuan akhir: {p.netloc}{p.path[:40]} param={sorted(q)}")

def save(data):
    open(OUT, "wb").write(data); print("xlsx valid, disimpan"); sys.exit(0)

# 1) API shares
tok = "u!" + base64.urlsafe_b64encode(URL.encode()).decode().rstrip("=")
d, _ = get(f"https://api.onedrive.com/v1.0/shares/{tok}/root/content", "Cara 1 (API shares)")
if d: save(d)

# 2) buka link (ikuti pengalihan) untuk tahu alamat akhir
d, final = get(URL, "Cara 2 (buka link)")
if d: save(d)
describe(final)

# 3) alamat akhir + download=1
d, _ = get(with_param(final, download="1"), "Cara 3 (akhir + download=1)")
if d: save(d)

# 4) pola onedrive.live.com/download?cid&resid&authkey
q = {k: v[0] for k, v in parse_qs(urlparse(final).query).items()}
rid = q.get("resid") or q.get("id")
if rid and q.get("authkey"):
    p = {"resid": rid, "authkey": q["authkey"]}
    if q.get("cid"): p["cid"] = q["cid"]
    d, _ = get("https://onedrive.live.com/download?" + urlencode(p), "Cara 4 (live.com/download)")
    if d: save(d)

# 5) API shares memakai alamat akhir
tok2 = "u!" + base64.urlsafe_b64encode(final.encode()).decode().rstrip("=")
d, _ = get(f"https://api.onedrive.com/v1.0/shares/{tok2}/root/content", "Cara 5 (API shares, alamat akhir)")
if d: save(d)

print("GAGAL: tidak ada cara yang berhasil mengunduh xlsx.")
sys.exit(1)
