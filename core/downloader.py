"""
core/downloader.py
Download XBRL ZIP files from IDX.
Strategi:
  - curl_cffi Session (TLS & cookies persistent, mirip browser asli)
  - sec-ch-ua headers matching Chrome version
  - Warm-up session lewat halaman IDX sebelum download
  - Retry cepat dengan ganti Chrome version
"""
import os
import time
import random

try:
    from curl_cffi import requests as cffi_requests
    HAS_CFFI = True
except ImportError:
    import requests as cffi_requests
    HAS_CFFI = False

MAX_RETRIES = 4

# Rotasi versi Chrome — tiap versi punya TLS fingerprint berbeda
CHROME_PROFILES = {
    "chrome110": {
        "impersonate": "chrome110",
        "sec-ch-ua": '"Chromium";v="110", "Not A(Brand";v="24", "Google Chrome";v="110"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
    },
    "chrome116": {
        "impersonate": "chrome116",
        "sec-ch-ua": '"Google Chrome";v="116", "Not)A;Brand";v="24", "Chromium";v="116"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
    },
    "chrome120": {
        "impersonate": "chrome120",
        "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
    },
    "chrome124": {
        "impersonate": "chrome124",
        "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
    },
}
CHROME_KEYS = list(CHROME_PROFILES.keys())

# Halaman untuk warm-up (dikunjungi lewat session sebelum download)
WARMUP_URLS = [
    "https://www.idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/",
    "https://www.idx.co.id/id/perusahaan-tercatat/daftar-saham/",
]


def _make_headers(profile: dict) -> dict:
    return {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "sec-ch-ua": profile["sec-ch-ua"],
        "sec-ch-ua-mobile": profile["sec-ch-ua-mobile"],
        "sec-ch-ua-platform": profile["sec-ch-ua-platform"],
        "Referer": "https://www.idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/",
        "Cache-Control": "max-age=0",
    }


def _create_session(chrome_key: str, log_callback=None) -> tuple:
    """
    Buat Session baru dan warm-up dengan mengunjungi halaman IDX.
    Return (session, chrome_key) yang sudah ter-warm-up.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    profile = CHROME_PROFILES[chrome_key]
    warmup_headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "sec-ch-ua": profile["sec-ch-ua"],
        "sec-ch-ua-mobile": profile["sec-ch-ua-mobile"],
        "sec-ch-ua-platform": profile["sec-ch-ua-platform"],
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0",
    }

    session = cffi_requests.Session(impersonate=chrome_key)
    warmup_url = random.choice(WARMUP_URLS)
    try:
        resp = session.get(warmup_url, headers=warmup_headers, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            log(f"  🔥 Session warm-up OK [{chrome_key}] — cookies: {len(session.cookies)} set")
        else:
            log(f"  ⚠ Warm-up [{resp.status_code}] — lanjut tanpa cookies")
    except Exception as e:
        log(f"  ⚠ Warm-up error: {str(e)[:60]}")
    
    time.sleep(random.uniform(0.5, 1.5))
    return session, chrome_key


def _download_single(session, url: str, save_path: str, chrome_key: str,
                     log_callback=None) -> str:
    """Download satu file via Session yang sudah warm-up."""
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    profile = CHROME_PROFILES[chrome_key]
    headers = _make_headers(profile)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=headers, timeout=60, allow_redirects=True)

            if resp.status_code == 200:
                content = resp.content
                if len(content) > 4 and content[:2] == b'PK':
                    with open(save_path, "wb") as f:
                        f.write(content)
                    return "success"
                else:
                    # Server return HTML bukan ZIP — mungkin session expired
                    if attempt < MAX_RETRIES:
                        log(f"  ⚠ Bukan ZIP (HTML response). Retry {attempt}...")
                        time.sleep(random.uniform(2, 4))
                        continue
                    return "not_zip"

            elif resp.status_code == 404:
                return "not_found"

            elif resp.status_code in (403, 429):
                if attempt < MAX_RETRIES:
                    # Wait singkat, ganti Chrome profile, lanjutkan dengan session yang ada
                    wait = 4 * attempt + random.uniform(1, 4)
                    new_key = random.choice([k for k in CHROME_KEYS if k != chrome_key])
                    log(f"  🔒 {resp.status_code} — Switch ke {new_key}, wait {wait:.0f}s...")
                    chrome_key = new_key
                    profile = CHROME_PROFILES[chrome_key]
                    headers = _make_headers(profile)
                    time.sleep(wait)
                else:
                    return "bot_detected"

            elif resp.status_code in (500, 502, 503, 504):
                if attempt < MAX_RETRIES:
                    wait = 3 * attempt + random.uniform(1, 3)
                    log(f"  ⚠ Server {resp.status_code} — retry {attempt} dalam {wait:.0f}s...")
                    time.sleep(wait)
                else:
                    return f"server_error"
            else:
                return f"failed_{resp.status_code}"

        except Exception as e:
            if attempt < MAX_RETRIES:
                wait = 3 * attempt + random.uniform(1, 3)
                log(f"  ⚠ Error: {str(e)[:60]} — retry {attempt} dalam {wait:.0f}s...")
                time.sleep(wait)
            else:
                return "error"

    return "bot_detected"


def download_all(links_file: str, output_folder: str,
                 progress_callback=None, log_callback=None,
                 stop_flag=None) -> dict:
    """
    Download semua ZIP dari companies.txt.
    Menggunakan persistent Session per batch untuk TLS continuity.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    stats = {"success": 0, "not_found": 0, "bot_detected": 0, "failed": 0, "skipped": 0}
    os.makedirs(output_folder, exist_ok=True)

    if not os.path.exists(links_file):
        log(f"Error: '{links_file}' tidak ditemukan. Jalankan Generate Links terlebih dahulu.")
        return stats

    with open(links_file, "r", encoding="utf-8") as f:
        lines = [l.strip() for l in f.readlines() if l.strip()]

    total = len(lines)
    log(f"Memproses {total} link download...")
    log(f"Mode: Persistent Session + sec-ch-ua matching + retry cepat\n")

    # Buat session awal
    chrome_key = random.choice(CHROME_KEYS)
    log(f"🔥 Inisialisasi session [{chrome_key}]...")
    session, chrome_key = _create_session(chrome_key, log)

    consecutive_bots = 0
    SESSION_REFRESH_INTERVAL = 60  # Refresh session setiap N download

    for idx, line in enumerate(lines):
        if stop_flag and stop_flag():
            log("Download dihentikan.")
            break

        parts = line.split(maxsplit=1)
        if len(parts) < 2:
            continue

        company_code, url = parts[0], parts[1]
        save_path = os.path.join(output_folder, f"{company_code}_inlineXBRL.zip")

        # Skip jika sudah ada
        if os.path.exists(save_path) and os.path.getsize(save_path) > 1000:
            log(f"[{idx+1}/{total}] {company_code} — ⏭ Skip (ada)")
            stats["skipped"] += 1
            if progress_callback: progress_callback(idx + 1, total)
            continue

        # Refresh session secara berkala atau setelah 3 bot berturut-turut
        if idx > 0 and (idx % SESSION_REFRESH_INTERVAL == 0 or consecutive_bots >= 3):
            new_key = random.choice(CHROME_KEYS)
            log(f"\n🔄 Refresh session [{new_key}]...")
            session.close()
            session, chrome_key = _create_session(new_key, log)
            consecutive_bots = 0

        log(f"[{idx+1}/{total}] ⬇ {company_code} [{chrome_key}]...")
        result = _download_single(session, url, save_path, chrome_key, log)

        if result == "success":
            log(f"  ✓ OK")
            stats["success"] += 1
            consecutive_bots = 0
            time.sleep(random.uniform(0.5, 1.5))

        elif result == "not_found":
            log(f"  — 404")
            stats["not_found"] += 1
            time.sleep(random.uniform(0.2, 0.6))

        elif result == "bot_detected":
            log(f"  🔒 Bot persisten — refresh session & tunggu 15s...")
            stats["bot_detected"] += 1
            consecutive_bots += 1
            session.close()
            new_key = random.choice(CHROME_KEYS)
            session, chrome_key = _create_session(new_key, log)
            time.sleep(15 + random.uniform(3, 8))

        else:
            log(f"  ✗ {result}")
            stats["failed"] += 1
            time.sleep(random.uniform(0.8, 2.0))

        if progress_callback: progress_callback(idx + 1, total)

    try:
        session.close()
    except Exception:
        pass

    log(f"\n{'='*40}")
    log(f"✓ Berhasil  : {stats['success']}")
    log(f"⏭ Skip      : {stats['skipped']}")
    log(f"— Tdk ada   : {stats['not_found']}")
    log(f"🔒 Bot block : {stats['bot_detected']}")
    log(f"✗ Gagal     : {stats['failed']}")
    return stats
