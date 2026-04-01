"""
core/link_generator.py
Generate download links for IDX financial reports (XBRL).
"""
import json
from datetime import datetime

try:
    from curl_cffi import requests
except ImportError:
    import requests

# URL template — {year} dan {period} dan {code} diganti secara dinamis
BASE_URL_TEMPLATE = (
    "https://www.idx.co.id"
    "/Portals/0/StaticData/ListedCompanies/Corporate_Actions/New_Info_JSX/"
    "Jenis_Informasi/01_Laporan_Keuangan/02_Soft_Copy_Laporan_Keuangan/"
    "/Laporan%20Keuangan%20Tahun%20{year}/{period}/{code}/inlineXBRL.zip"
)

PERIOD_MAP = {
    "Q1 (TW1)": "TW1",
    "Q2 (TW2)": "TW2",
    "Q3 (TW3)": "TW3",
    "Tahunan (Audit)": "Audit",
}


def get_period_code(period_label: str) -> str:
    """Konversi label periode tampilan ke kode URL."""
    return PERIOD_MAP.get(period_label, period_label)


def build_url(year: int, period_code: str, ticker: str) -> str:
    return BASE_URL_TEMPLATE.format(year=year, period=period_code, code=ticker)


def get_all_tickers(log_callback=None) -> list:
    """Ambil semua ticker perusahaan dari API IDX."""
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("Mengambil daftar emiten dari IDX...")

    endpoints = [
        "https://www.idx.co.id/primary/StockData/GetSecuritiesData?emitenType=s&start=0&length=9999&sortBy=c&sortType=asc",
        "https://www.idx.co.id/primary/ListedCompany/GetCompanyProfiles?emitenType=s&start=0&length=9999",
    ]

    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.idx.co.id/en-us/market-data/shares/",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.idx.co.id",
    }

    response = None
    for url in endpoints:
        try:
            resp = requests.get(url, impersonate="chrome120", headers=headers, timeout=30)
            if resp.status_code == 200:
                response = resp
                break
            log(f"  [{resp.status_code}] {url}")
        except Exception as e:
            log(f"  Error: {e}")

    if response is None:
        log("Gagal mengambil data ticker dari semua endpoint.")
        return []

    try:
        data = response.json()
        if "data" in data:
            tickers = [item["KodeEmiten"] for item in data["data"]]
            tickers = [t for t in tickers if isinstance(t, str) and t.strip()]
            tickers = sorted(list(set(tickers)))
            return tickers
        else:
            log(f"Struktur JSON tidak terduga. Keys: {list(data.keys())}")
            return []
    except Exception as e:
        log(f"Error parsing response: {e}")
        return []


def generate_links(year: int, period_label: str, output_path: str,
                   progress_callback=None, log_callback=None) -> int:
    """
    Generate file companies.txt berisi kode emiten + URL download.

    Returns:
        Jumlah ticker yang berhasil digenerate, atau -1 jika gagal.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    period_code = get_period_code(period_label)
    log(f"Tahun: {year} | Periode: {period_label} ({period_code})")

    tickers = get_all_tickers(log_callback=log_callback)
    if not tickers:
        log("Tidak ada ticker yang ditemukan. Proses dihentikan.")
        return -1

    log(f"Ditemukan {len(tickers)} perusahaan.")

    with open(output_path, "w", encoding="utf-8") as f:
        for i, code in enumerate(tickers):
            url = build_url(year, period_code, code)
            f.write(f"{code} {url}\n")
            if progress_callback:
                progress_callback(i + 1, len(tickers))

    log(f"Link berhasil disimpan ke: {output_path}")
    return len(tickers)


if __name__ == "__main__":
    year = datetime.now().year
    generate_links(year, "Tahunan (Audit)", "companies.txt")
