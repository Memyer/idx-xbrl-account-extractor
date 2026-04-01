"""
gui/app.py — IDX Superapp (CustomTkinter)
Includes: Settings, Account Selector (search + categories), Run Pipeline, Log, About
"""
import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import threading
from datetime import datetime
import re
import os, sys, shutil

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.link_generator import generate_links, PERIOD_MAP
from core.downloader import download_all
from core.unzipper import unzip_all
from core.xbrl_extractor import (
    extract_all, ALL_METRICS, DEFAULT_METRICS
)
from core.data_cleaner import clean_data
from core.amount_normalizer import normalize_to_full_idr, ROUNDING_COLUMN, CURRENCY_COLUMN
from core.csv_exporter import export_to_excel

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C_BG      = "#0f1117"
C_SURFACE = "#1a1d2e"
C_CARD    = "#21253a"
C_BORDER  = "#2e3354"
C_ACCENT  = "#4f8ef7"
C_SUCCESS = "#2dca72"
C_WARNING = "#f5a623"
C_DANGER  = "#e74c3c"
C_TEXT    = "#e8eaf6"
C_MUTED   = "#8892b0"
C_WHITE   = "#ffffff"
C_CAT     = {
    "Informasi Umum":          "#a78bfa",
    "Neraca (Balance Sheet)":  "#34d399",
    "Laba Rugi (Income Statement)": "#60a5fa",
    "Arus Kas (Cash Flow)":    "#fb923c",
}

NORMALIZATION_REFERENCE_METRICS = [ROUNDING_COLUMN, CURRENCY_COLUMN]


class IDXSuperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("IDX Superapp")
        self.geometry("1100x760")
        self.minsize(1000, 680)
        self.configure(fg_color=C_BG)

        self._stop_flag = False
        self._running   = False

        # ── Settings vars ──
        self.var_year     = tk.IntVar(value=datetime.now().year)
        self.var_period   = tk.StringVar(value="Tahunan (Audit)")
        self.var_base_dir = tk.StringVar(value=os.getcwd())
        self.var_usd_rate = tk.StringVar(value="16000")
        self.var_clean_keuangan = tk.BooleanVar(value=True)
        self.var_clean_bunga    = tk.BooleanVar(value=True)

        self.var_companies_file = tk.StringVar()
        self.var_download_dir   = tk.StringVar()
        self.var_xbrl_dir       = tk.StringVar()
        self.var_extract_dir    = tk.StringVar()
        self._update_paths()
        self.var_base_dir.trace_add("write", lambda *_: self._update_paths())

        # ── Account selector state ──
        self._all_account_options = self._load_accounts_from_file() or list(ALL_METRICS)
        self._selected_metrics = [m for m in DEFAULT_METRICS if m in self._all_account_options]
        if not self._selected_metrics and self._all_account_options:
            self._selected_metrics = self._all_account_options[:min(20, len(self._all_account_options))]

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter_metrics())
        self._filtered_metrics = list(self._all_account_options)

        # UI refs for account selector
        self.account_listbox = None
        self.selected_list = None
        self.lbl_selected_count = None

        self._build_ui()

    def _load_accounts_from_file(self) -> list:
        """Load Indonesian account names from akun_indonesia.txt."""
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        candidates = [
            os.path.join(project_root, "akun_indonesia.txt"),
            os.path.join(app_root, "akun_indonesia.txt"),
        ]

        account_names = []
        seen = set()
        for path in candidates:
            if not os.path.exists(path):
                continue
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        m = re.match(r"^\s*\d+\.\s+(.+?)\s*$", line)
                        if not m:
                            continue
                        name = m.group(1).strip()
                        if not name or name in seen:
                            continue
                        seen.add(name)
                        account_names.append(name)
            except Exception:
                continue
            if account_names:
                break
        return account_names

    def _update_paths(self):
        b = self.var_base_dir.get()
        self.var_companies_file.set(os.path.join(b, "companies.txt"))
        self.var_download_dir.set(os.path.join(b, "FinancialStatements"))
        self.var_xbrl_dir.set(os.path.join(b, "Folder_XBRL"))
        self.var_extract_dir.set(os.path.join(b, "ExtractedData_XBRL"))

    # ══════════════════════════════════════
    #  BUILD UI
    # ══════════════════════════════════════
    def _build_ui(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color="#12162b", corner_radius=0, height=58)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="📊  IDX Superapp",
                     font=ctk.CTkFont("Segoe UI", 17, "bold"),
                     text_color=C_WHITE).pack(side="left", padx=20, pady=14)
        ctk.CTkLabel(hdr, text="Otomasi Ekstraksi Data Keuangan XBRL — Bursa Efek Indonesia",
                     font=ctk.CTkFont("Segoe UI", 10), text_color=C_MUTED).pack(side="left")

        # Tabs
        self.tabs = ctk.CTkTabview(
            self, fg_color=C_SURFACE,
            segmented_button_fg_color="#12162b",
            segmented_button_selected_color=C_ACCENT,
            segmented_button_selected_hover_color="#3a7de8",
            segmented_button_unselected_color="#12162b",
            segmented_button_unselected_hover_color="#1e2240",
            border_color=C_BORDER, border_width=1, corner_radius=12
        )
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        for tab in ["⚙️ Pengaturan", "📋 Pilih Akun", "🚀 Jalankan", "📃 Log", "ℹ️ Tentang"]:
            self.tabs.add(tab)

        self._build_settings()
        self._build_account_selector()
        self._build_run()
        self._build_log()
        self._build_about()

    # ══════════════════════════════════════
    #  TAB: PENGATURAN
    # ══════════════════════════════════════
    def _build_settings(self):
        tab = self.tabs.tab("⚙️ Pengaturan")
        scroll = ctk.CTkScrollableFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Periode
        self._section(scroll, "📅  Konfigurasi Periode")
        c1 = self._card(scroll)
        row_yr = ctk.CTkFrame(c1, fg_color="transparent")
        row_yr.pack(fill="x", pady=6)
        ctk.CTkLabel(row_yr, text="Tahun Laporan", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_TEXT, width=150, anchor="w").pack(side="left")
        ctk.CTkButton(row_yr, text="◀", width=32, height=32, fg_color=C_CARD,
                      hover_color=C_BORDER,
                      command=lambda: self.var_year.set(self.var_year.get()-1)).pack(side="left", padx=(0,4))
        ctk.CTkLabel(row_yr, textvariable=self.var_year,
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C_ACCENT, width=60).pack(side="left")
        ctk.CTkButton(row_yr, text="▶", width=32, height=32, fg_color=C_CARD,
                      hover_color=C_BORDER,
                      command=lambda: self.var_year.set(self.var_year.get()+1)).pack(side="left", padx=(4,0))

        row_pd = ctk.CTkFrame(c1, fg_color="transparent")
        row_pd.pack(fill="x", pady=6)
        ctk.CTkLabel(row_pd, text="Periode", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_TEXT, width=150, anchor="w").pack(side="left")
        ctk.CTkComboBox(row_pd, variable=self.var_period, values=list(PERIOD_MAP.keys()),
                        state="readonly", width=220,
                        font=ctk.CTkFont("Segoe UI", 11),
                        dropdown_font=ctk.CTkFont("Segoe UI", 11)).pack(side="left")
        ctk.CTkLabel(c1, text="  Q1=TW1 (Mar) · Q2=TW2 (Jun) · Q3=TW3 (Sep) · Tahunan=Audit (Des)",
                     font=ctk.CTkFont("Segoe UI", 9), text_color=C_MUTED).pack(anchor="w", pady=(0, 4))

        # Folder
        self._section(scroll, "📁  Folder Output")
        c2 = self._card(scroll)
        row_dir = ctk.CTkFrame(c2, fg_color="transparent")
        row_dir.pack(fill="x", pady=4)
        ctk.CTkLabel(row_dir, text="Folder Kerja", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_TEXT, width=120, anchor="w").pack(side="left")
        ctk.CTkEntry(row_dir, textvariable=self.var_base_dir,
                     font=ctk.CTkFont("Consolas", 10), fg_color=C_CARD,
                     border_color=C_BORDER).pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(row_dir, text="Browse", width=80, fg_color=C_BORDER,
                      hover_color="#3a3d6a", command=self._browse).pack(side="left")
        for lbl, var, desc in [
            ("companies.txt",       self.var_companies_file, "Link download emiten"),
            ("FinancialStatements/", self.var_download_dir,  "ZIP hasil download"),
            ("Folder_XBRL/",        self.var_xbrl_dir,       "XBRL extracted"),
            ("ExtractedData_XBRL/", self.var_extract_dir,    "Output CSV + Excel"),
        ]:
            r = ctk.CTkFrame(c2, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text=f"  ↳ {lbl}", font=ctk.CTkFont("Consolas", 9),
                         text_color=C_ACCENT, width=200, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=desc, font=ctk.CTkFont("Segoe UI", 9),
                         text_color=C_MUTED).pack(side="left", padx=8)

        # Opsi cleaning
        self._section(scroll, "🧹  Opsi Pembersihan Nilai")
        c3 = self._card(scroll)
        ctk.CTkLabel(c3, text="Ubah nilai negatif ke positif untuk kolom:",
                     font=ctk.CTkFont("Segoe UI", 10), text_color=C_MUTED).pack(anchor="w", pady=(0, 8))
        ctk.CTkCheckBox(c3, text="Beban bunga dan keuangan",
                        variable=self.var_clean_keuangan,
                        font=ctk.CTkFont("Segoe UI", 11),
                        fg_color=C_ACCENT, hover_color="#3a7de8",
                        checkbox_width=20, checkbox_height=20).pack(anchor="w", pady=3)
        ctk.CTkCheckBox(c3, text="Beban bunga",
                        variable=self.var_clean_bunga,
                        font=ctk.CTkFont("Segoe UI", 11),
                        fg_color=C_ACCENT, hover_color="#3a7de8",
                        checkbox_width=20, checkbox_height=20).pack(anchor="w", pady=3)

        # Opsi normalisasi IDR
        self._section(scroll, "💱  Opsi Normalisasi IDR")
        c4 = self._card(scroll)
        ctk.CTkLabel(
            c4,
            text=(
                "Tahap 7 akan mengubah nilai akun terpilih menjadi Full Amount IDR\n"
                "berdasarkan kolom pembulatan dan mata uang pelaporan."
            ),
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=C_MUTED,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))
        usd_row = ctk.CTkFrame(c4, fg_color="transparent")
        usd_row.pack(fill="x")
        ctk.CTkLabel(usd_row, text="Kurs USD ke IDR",
                     font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_TEXT, width=150, anchor="w").pack(side="left")
        ctk.CTkEntry(usd_row, textvariable=self.var_usd_rate,
                     width=180, font=ctk.CTkFont("Consolas", 11),
                     fg_color=C_CARD, border_color=C_BORDER).pack(side="left")
        ctk.CTkLabel(usd_row, text="contoh: 16000",
                     font=ctk.CTkFont("Segoe UI", 9),
                     text_color=C_MUTED).pack(side="left", padx=8)

    def _browse(self):
        d = filedialog.askdirectory(title="Pilih Folder Output")
        if d:
            self.var_base_dir.set(d.replace("/", "\\"))

    # ══════════════════════════════════════
    #  TAB: PILIH AKUN (ACCOUNT SELECTOR)
    # ══════════════════════════════════════
    def _build_account_selector(self):
        tab = self.tabs.tab("📋 Pilih Akun")

        # --- Top bar: search + quick buttons ---
        top = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        top.pack(fill="x", padx=12, pady=(10, 0))

        search_row = ctk.CTkFrame(top, fg_color="transparent")
        search_row.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(search_row, text="🔍", font=ctk.CTkFont("Segoe UI", 14),
                     text_color=C_MUTED).pack(side="left", padx=(0, 6))
        ctk.CTkEntry(search_row, textvariable=self._search_var,
                     placeholder_text="Cari akun... (contoh: aset, laba, kas)",
                     font=ctk.CTkFont("Segoe UI", 11),
                     fg_color=C_CARD, border_color=C_BORDER, height=36).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(search_row, text="✕ Clear", width=70, height=36,
                      fg_color=C_CARD, hover_color=C_BORDER,
                      font=ctk.CTkFont("Segoe UI", 10),
                      command=lambda: self._search_var.set("")).pack(side="left", padx=(6, 0))

        picker_wrap = ctk.CTkFrame(top, fg_color=C_CARD, corner_radius=8,
                                   border_color=C_BORDER, border_width=1)
        picker_wrap.pack(fill="x", pady=(0, 8))

        picker_hdr = ctk.CTkFrame(picker_wrap, fg_color="transparent")
        picker_hdr.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(picker_hdr, text="Hasil pencarian akun (scroll untuk naik/turun)",
                     font=ctk.CTkFont("Segoe UI", 10, "bold"),
                     text_color=C_TEXT).pack(side="left")
        ctk.CTkButton(picker_hdr, text="+ Tambah Pilihan", width=130, height=30,
                      fg_color=C_ACCENT, hover_color="#3a7de8",
                      font=ctk.CTkFont("Segoe UI", 9, "bold"),
                      command=self._add_selected_metric).pack(side="right")

        lb_wrap = ctk.CTkFrame(picker_wrap, fg_color="#191d30", corner_radius=6)
        lb_wrap.pack(fill="x", padx=10, pady=(0, 10))

        self.account_listbox = tk.Listbox(
            lb_wrap,
            height=7,
            selectmode=tk.SINGLE,
            bg="#191d30",
            fg=C_TEXT,
            selectbackground=C_ACCENT,
            selectforeground=C_WHITE,
            highlightthickness=0,
            relief="flat",
            activestyle="none",
            font=("Segoe UI", 10),
        )
        self.account_listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)

        lb_scroll = tk.Scrollbar(lb_wrap, orient="vertical", command=self.account_listbox.yview)
        lb_scroll.pack(side="right", fill="y", padx=(6, 8), pady=8)
        self.account_listbox.configure(yscrollcommand=lb_scroll.set)

        self.account_listbox.bind("<Double-Button-1>", lambda _e: self._add_selected_metric())
        self.account_listbox.bind("<MouseWheel>", self._on_account_list_scroll)
        self.account_listbox.bind("<Button-4>", self._on_account_list_scroll)
        self.account_listbox.bind("<Button-5>", self._on_account_list_scroll)

        stats_row = ctk.CTkFrame(top, fg_color="transparent")
        stats_row.pack(fill="x", pady=(0, 8))

        self.lbl_selected_count = ctk.CTkLabel(
            stats_row, text=self._count_text(),
            font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=C_ACCENT)
        self.lbl_selected_count.pack(side="left")

        btn_kw = dict(height=28, font=ctk.CTkFont("Segoe UI", 9, "bold"),
                      corner_radius=6, fg_color=C_CARD, hover_color=C_BORDER)
        ctk.CTkButton(stats_row, text="☑ Pilih Semua", width=110, **btn_kw,
                      command=self._select_all).pack(side="right", padx=(4, 0))
        ctk.CTkButton(stats_row, text="☐ Hapus Semua", width=110, **btn_kw,
                      command=self._deselect_all).pack(side="right", padx=(4, 0))
        ctk.CTkButton(stats_row, text="↺ Default", width=80, **btn_kw,
                      command=self._reset_to_default).pack(side="right", padx=(4, 0))

        ctk.CTkFrame(tab, fg_color=C_BORDER, height=1).pack(fill="x", padx=12)

        # --- Selected accounts list ---
        selected_wrap = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        selected_wrap.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        ctk.CTkLabel(selected_wrap, text="Akun terpilih", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_WHITE).pack(anchor="w", pady=(0, 6))

        self.selected_list = ctk.CTkScrollableFrame(
            selected_wrap, fg_color=C_CARD, corner_radius=8,
            scrollbar_button_color=C_BORDER,
        )
        self.selected_list.pack(fill="both", expand=True)

        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _count_text(self) -> str:
        return f"{len(self._selected_metrics)} dari {len(self._all_account_options)} akun dipilih"

    def _refresh_dropdown_options(self):
        if not self.account_listbox:
            return
        q = self._search_var.get().strip().lower()
        selected = set(self._selected_metrics)
        self._filtered_metrics = [
            m for m in self._all_account_options
            if m not in selected and (not q or q in m.lower())
        ]

        self.account_listbox.delete(0, tk.END)
        if self._filtered_metrics:
            for metric in self._filtered_metrics:
                self.account_listbox.insert(tk.END, metric)
            self.account_listbox.selection_set(0)
        else:
            self.account_listbox.insert(tk.END, "(Tidak ada hasil)")

    def _refresh_selected_metrics_view(self):
        if not self.selected_list:
            return
        for child in self.selected_list.winfo_children():
            child.destroy()

        if not self._selected_metrics:
            ctk.CTkLabel(self.selected_list,
                         text="Belum ada akun dipilih.",
                         font=ctk.CTkFont("Segoe UI", 10),
                         text_color=C_MUTED).pack(anchor="w", padx=8, pady=8)
        else:
            for metric in self._selected_metrics:
                row = ctk.CTkFrame(self.selected_list, fg_color="#191d30", corner_radius=6)
                row.pack(fill="x", padx=6, pady=3)
                ctk.CTkLabel(row, text=metric,
                             font=ctk.CTkFont("Segoe UI", 10),
                             text_color=C_TEXT, anchor="w").pack(side="left", fill="x", expand=True, padx=10, pady=6)
                ctk.CTkButton(row, text="Hapus", width=70, height=26,
                              fg_color=C_DANGER, hover_color="#c0392b",
                              font=ctk.CTkFont("Segoe UI", 9, "bold"),
                              command=lambda m=metric: self._remove_selected_metric(m)).pack(side="right", padx=8)

        self.lbl_selected_count.configure(text=self._count_text())

    def _add_selected_metric(self):
        if not self.account_listbox:
            return
        cur = self.account_listbox.curselection()
        if not cur:
            return
        metric = self.account_listbox.get(cur[0]).strip()
        if not metric or metric == "(Tidak ada hasil)":
            return
        if metric not in self._selected_metrics:
            self._selected_metrics.append(metric)
        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _on_account_list_scroll(self, event):
        if not self.account_listbox:
            return "break"
        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        if delta != 0:
            self.account_listbox.yview_scroll(delta, "units")
        return "break"

    def _remove_selected_metric(self, metric: str):
        self._selected_metrics = [m for m in self._selected_metrics if m != metric]
        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _select_all(self):
        self._selected_metrics = list(self._all_account_options)
        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _deselect_all(self):
        self._selected_metrics = []
        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _reset_to_default(self):
        self._selected_metrics = [m for m in DEFAULT_METRICS if m in self._all_account_options]
        self._refresh_dropdown_options()
        self._refresh_selected_metrics_view()

    def _filter_metrics(self):
        self._refresh_dropdown_options()

    def get_selected_metrics(self) -> list:
        """Return list akun yang user pilih."""
        return list(self._selected_metrics)

    # ══════════════════════════════════════
    #  TAB: JALANKAN
    # ══════════════════════════════════════
    def _build_run(self):
        tab = self.tabs.tab("🚀 Jalankan")
        outer = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        outer.pack(fill="both", expand=True, padx=16, pady=12)

        self._section(outer, "🔄  Pipeline Ekstraksi Data")

        self.steps = [
            ("generate", "1. Generate Links",        "🔗", "Ambil semua ticker IDX & buat companies.txt"),
            ("download",  "2. Download XBRL ZIP",    "⬇️", "Download inlineXBRL.zip dengan session anti-bot"),
            ("unzip",     "3. Ekstrak ZIP",           "📦", "Ekstrak ZIP ke Folder_XBRL"),
            ("extract",   "4. Ekstrak Data XBRL",    "🔍", "Parse HTML, ekstrak akun yang dipilih → CSV"),
            ("clean",     "5. Bersihkan Data",        "🧹", "Ubah nilai negatif sesuai pengaturan"),
            ("normalize", "6. Normalisasi Full IDR",  "💱", "Konversi nilai sesuai pembulatan + mata uang"),
            ("export",    "7. Export CSV + Excel",    "📊", "Simpan 2 set output: sebelum & sesudah konversi"),
        ]
        self._step_w = {}
        for key, label, icon, desc in self.steps:
            self._make_step_row(outer, key, label, icon, desc)

        btn_row = ctk.CTkFrame(outer, fg_color="transparent")
        btn_row.pack(fill="x", pady=(12, 0))
        self.btn_run_all = ctk.CTkButton(btn_row, text="▶  Jalankan Semua",
                                                        font=ctk.CTkFont("Segoe UI", 18, "bold"),
                                          fg_color=C_SUCCESS, hover_color="#25a85e",
                                                        height=60, width=500, corner_radius=12,
                                          command=self._run_all)
        self.btn_run_all.pack(side="left", padx=(0, 10))
        self.btn_stop = ctk.CTkButton(btn_row, text="⏹  Stop",
                                                    font=ctk.CTkFont("Segoe UI", 16, "bold"),
                                       fg_color=C_DANGER, hover_color="#c0392b",
                                                    height=60, width=300, corner_radius=12,
                                       state="disabled", command=self._stop_pipeline)
        self.btn_stop.pack(side="left")

        ctk.CTkLabel(outer, text="Progress keseluruhan:",
                     font=ctk.CTkFont("Segoe UI", 9), text_color=C_MUTED).pack(anchor="w", pady=(14, 4))
        self.overall_pb = ctk.CTkProgressBar(outer, height=10, corner_radius=5,
                                              fg_color=C_CARD, progress_color=C_ACCENT)
        self.overall_pb.set(0)
        self.overall_pb.pack(fill="x")
        self.lbl_overall = ctk.CTkLabel(outer, text="Siap",
                                         font=ctk.CTkFont("Segoe UI", 9), text_color=C_MUTED)
        self.lbl_overall.pack(anchor="w", pady=(4, 0))

    def _make_step_row(self, parent, key, label, icon, desc):
        card = ctk.CTkFrame(parent, fg_color=C_CARD, corner_radius=8)
        card.pack(fill="x", pady=3)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=8)
        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(left, text=f"{icon}  {label}",
                     font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_TEXT, anchor="w").pack(anchor="w")
        ctk.CTkLabel(left, text=desc, font=ctk.CTkFont("Segoe UI", 9),
                     text_color=C_MUTED, anchor="w").pack(anchor="w")
        right = ctk.CTkFrame(inner, fg_color="transparent")
        right.pack(side="right", padx=(12, 0))
        pb = ctk.CTkProgressBar(right, width=160, height=8, corner_radius=4,
                                 fg_color=C_BORDER, progress_color=C_ACCENT)
        pb.set(0)
        pb.pack(side="left", padx=(0, 10))
        lbl = ctk.CTkLabel(right, text="Menunggu",
                            font=ctk.CTkFont("Segoe UI", 9),
                            text_color=C_MUTED, width=90, anchor="w")
        lbl.pack(side="left", padx=(0, 8))
        btn = ctk.CTkButton(right, text="▶ Run", width=80,
                             font=ctk.CTkFont("Segoe UI", 10, "bold"),
                             fg_color=C_ACCENT, hover_color="#3a7de8",
                             height=28, corner_radius=6,
                             command=lambda k=key: self._run_step(k))
        btn.pack(side="left")
        self._step_w[key] = {"pb": pb, "lbl": lbl, "btn": btn}

    # ══════════════════════════════════════
    #  TAB: LOG
    # ══════════════════════════════════════
    def _build_log(self):
        tab = self.tabs.tab("📃 Log")
        outer = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        outer.pack(fill="both", expand=True, padx=12, pady=12)
        top = ctk.CTkFrame(outer, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(top, text="📃  Log Output",
                     font=ctk.CTkFont("Segoe UI", 12, "bold"),
                     text_color=C_WHITE).pack(side="left")
        ctk.CTkButton(top, text="🗑 Hapus Log", width=110, height=30,
                      fg_color=C_CARD, hover_color=C_BORDER,
                      font=ctk.CTkFont("Segoe UI", 10),
                      command=self._clear_log).pack(side="right")
        self.log_box = ctk.CTkTextbox(
            outer, fg_color="#090d1a", text_color="#58a6ff",
            font=ctk.CTkFont("Consolas", 10), wrap="word",
            corner_radius=8, border_color=C_BORDER, border_width=1,
            scrollbar_button_color=C_BORDER, state="disabled"
        )
        self.log_box.pack(fill="both", expand=True)
        tb = self.log_box._textbox
        tb.tag_configure("ok",   foreground=C_SUCCESS)
        tb.tag_configure("err",  foreground=C_DANGER)
        tb.tag_configure("warn", foreground=C_WARNING)
        tb.tag_configure("info", foreground=C_ACCENT)
        tb.tag_configure("dim",  foreground=C_MUTED)

    # ══════════════════════════════════════
    #  TAB: TENTANG
    # ══════════════════════════════════════
    def _build_about(self):
        tab = self.tabs.tab("ℹ️ Tentang")
        outer = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=0)
        outer.pack(fill="both", expand=True, padx=40, pady=40)
        ctk.CTkLabel(outer, text="📊  IDX Superapp",
                     font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=C_WHITE).pack(pady=(0, 6))
        ctk.CTkLabel(outer, text="Ekstraksi Data Keuangan Otomatis dari IDX (Inline XBRL)",
                     font=ctk.CTkFont("Segoe UI", 11), text_color=C_MUTED).pack()
        ctk.CTkFrame(outer, fg_color=C_BORDER, height=2).pack(fill="x", pady=24)
        rows = [
            ("Pipeline",    "Generate → Download → Unzip → Ekstrak XBRL → Clean → Normalisasi IDR → Export CSV+Excel"),
            ("Akun XBRL",  f"{len(ALL_METRICS)} akun dari 4 kategori — dapat dipilih per akun"),
            ("Anti-bot",   "Persistent Session + sec-ch-ua matching + retry cepat"),
            ("Output",     "CSV + Excel (XLSX) otomatis di ExtractedData_XBRL/"),
            ("Sumber",     "idx.co.id — Inline XBRL Financial Statements"),
        ]
        for lbl, val in rows:
            r = ctk.CTkFrame(outer, fg_color="transparent")
            r.pack(fill="x", pady=5)
            ctk.CTkLabel(r, text=f"{lbl}:", font=ctk.CTkFont("Segoe UI", 10, "bold"),
                         text_color=C_ACCENT, width=130, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=val, font=ctk.CTkFont("Segoe UI", 10),
                         text_color=C_TEXT, wraplength=680, justify="left").pack(side="left")

    # ══════════════════════════════════════
    #  HELPERS
    # ══════════════════════════════════════
    def _section(self, parent, text):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C_WHITE).pack(anchor="w", pady=(10, 4))

    def _card(self, parent):
        frame = ctk.CTkFrame(parent, fg_color=C_CARD, corner_radius=10,
                              border_color=C_BORDER, border_width=1)
        frame.pack(fill="x", pady=(2, 12), padx=2)
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)
        return inner

    def log(self, msg: str, tag: str = None):
        def _do():
            self.log_box.configure(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            tb = self.log_box._textbox
            tb.insert("end", f"[{ts}] {msg}\n", tag or "")
            tb.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("0.0", "end")
        self.log_box.configure(state="disabled")

    def _set_step(self, key, status, color=None, pb_val=None):
        def _do():
            w = self._step_w.get(key)
            if not w: return
            w["lbl"].configure(text=status, text_color=color or C_MUTED)
            if pb_val is not None: w["pb"].set(pb_val)
        self.after(0, _do)

    def _set_progress(self, key, cur, tot):
        def _do():
            w = self._step_w.get(key)
            if w and tot > 0: w["pb"].set(cur / tot)
        self.after(0, _do)

    def _set_overall(self, cur, tot, label=""):
        def _do():
            if tot > 0: self.overall_pb.set(cur / tot)
            if label: self.lbl_overall.configure(text=label)
        self.after(0, _do)

    def _set_btns(self, running: bool):
        def _do():
            s = "disabled" if running else "normal"
            self.btn_run_all.configure(state=s)
            self.btn_stop.configure(state="normal" if running else "disabled")
            for w in self._step_w.values():
                w["btn"].configure(state=s)
        self.after(0, _do)

    def _stop_pipeline(self):
        self._stop_flag = True
        self.log("⏹ Stop diminta...", "warn")

    def _parse_usd_rate(self) -> float:
        raw = self.var_usd_rate.get().strip().replace(" ", "")
        if not raw:
            raise ValueError("Kurs USD kosong. Isi kurs USD ke IDR terlebih dahulu.")

        s = raw
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            if s.count(",") == 1 and len(s.split(",")[1]) <= 2:
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "." in s:
            if s.count(".") > 1:
                s = s.replace(".", "")
            else:
                left, right = s.split(".")
                if left.isdigit() and right.isdigit() and len(right) == 3:
                    s = left + right

        rate = float(s)
        if rate <= 0:
            raise ValueError("Kurs USD harus lebih besar dari 0.")
        return rate

    # ══════════════════════════════════════
    #  PIPELINE LOGIC
    # ══════════════════════════════════════
    def _run_all(self):
        if self._running: return
        selected = self.get_selected_metrics()
        if not selected:
            messagebox.showwarning("Perhatian", "Belum ada akun yang dipilih!\nBuka tab 'Pilih Akun' untuk memilih.")
            return
        self._stop_flag = False
        self._running = True
        self._set_btns(True)
        self.log("=" * 55, "dim")
        self.log(f"▶  Pipeline dimulai — {len(selected)} akun dipilih", "info")

        def _go():
            keys = ["generate", "download", "unzip", "extract", "clean", "normalize", "export"]
            for i, key in enumerate(keys):
                if self._stop_flag: break
                self._set_overall(i, len(keys), f"Step {i+1}/{len(keys)} — {key}")
                self._do_step(key)
            self._set_overall(1, 1, "✓ Selesai!")
            self._running = False
            self._set_btns(False)
            if not self._stop_flag:
                self.log("✅  Semua selesai!", "ok")
                self.after(0, lambda: messagebox.showinfo(
                    "Selesai", "Pipeline berhasil!\n\nHasil tersimpan di ExtractedData_XBRL/\n"
                               "  • 00_SUMMARY_all_companies_raw.csv\n"
                               "  • 00_SUMMARY_all_companies_raw.xlsx\n"
                               "  • 00_SUMMARY_all_companies_full_idr.csv\n"
                               "  • 00_SUMMARY_all_companies_full_idr.xlsx"))
        threading.Thread(target=_go, daemon=True).start()

    def _run_step(self, key: str):
        if self._running: return
        self._stop_flag = False
        self._running = True
        self._set_btns(True)
        def _go():
            self._do_step(key)
            self._running = False
            self._set_btns(False)
        threading.Thread(target=_go, daemon=True).start()

    def _do_step(self, key: str):
        year   = self.var_year.get()
        period = self.var_period.get()
        cf     = self.var_companies_file.get()
        dl     = self.var_download_dir.get()
        xbrl   = self.var_xbrl_dir.get()
        ext    = self.var_extract_dir.get()
        csv_s  = os.path.join(ext, "00_SUMMARY_all_companies.csv")
        csv_raw = os.path.join(ext, "00_SUMMARY_all_companies_raw.csv")
        csv_full_idr = os.path.join(ext, "00_SUMMARY_all_companies_full_idr.csv")

        slog = lambda m: self.log(m)
        spg  = lambda c, t: self._set_progress(key, c, t)
        stop = lambda: self._stop_flag

        self._set_step(key, "Berjalan...", C_WARNING)
        try:
            if key == "generate":
                self.log(f"🔗 Generate — Tahun {year}, {period}", "info")
                n = generate_links(year, period, cf, progress_callback=spg, log_callback=slog)
                if n and n > 0:
                    self._set_step(key, f"✓ {n} link", C_SUCCESS, 1.0)
                    self.log(f"✓ {n} link → {cf}", "ok")
                else:
                    self._set_step(key, "Gagal", C_DANGER)

            elif key == "download":
                self.log(f"⬇️ Download...", "info")
                st = download_all(cf, dl, spg, slog, stop)
                self._set_step(key, f"✓ {st['success']} OK", C_SUCCESS, 1.0)
                self.log(f"✓ {st['success']} ✓ · {st['skipped']} skip · {st['bot_detected']} bot", "ok")

            elif key == "unzip":
                self.log("📦 Ekstrak ZIP...", "info")
                st = unzip_all(dl, xbrl, spg, slog, stop)
                self._set_step(key, f"✓ {st['success']}", C_SUCCESS, 1.0)

            elif key == "extract":
                selected = self.get_selected_metrics()
                if not selected:
                    self.log("⚠ Tidak ada akun dipilih! Buka tab Pilih Akun.", "warn")
                    self._set_step(key, "Tidak ada akun", C_WARNING)
                    return
                selected_for_extract = list(selected)
                added_refs = []
                for ref_metric in NORMALIZATION_REFERENCE_METRICS:
                    if ref_metric not in selected_for_extract:
                        selected_for_extract.append(ref_metric)
                        added_refs.append(ref_metric)
                if added_refs:
                    self.log("ℹ Tambah akun referensi normalisasi: " + ", ".join(added_refs), "dim")

                self.log(f"🔍 Ekstrak {len(selected_for_extract)} akun dari XBRL...", "info")
                r = extract_all(xbrl, ext, selected_metrics=selected_for_extract,
                                progress_callback=spg, log_callback=slog, stop_flag=stop)
                if r:
                    self._set_step(key, "✓ CSV dibuat", C_SUCCESS, 1.0)
                    self.log(f"✓ CSV → {r}", "ok")
                else:
                    self._set_step(key, "Gagal/kosong", C_WARNING)

            elif key == "clean":
                cols = []
                if self.var_clean_keuangan.get(): cols.append("Beban bunga dan keuangan")
                if self.var_clean_bunga.get(): cols.append("Beban bunga")
                if cols:
                    self.log(f"🧹 Bersihkan: {', '.join(cols)}", "info")
                else:
                    self.log("🧹 Tidak ada kolom dipilih — dilewati", "warn")
                ok = clean_data(csv_s, columns_to_clean=cols, log_callback=slog)
                self._set_step(key, "✓ Selesai" if ok else "Gagal",
                               C_SUCCESS if ok else C_DANGER, 1.0)

            elif key == "normalize":
                selected = self.get_selected_metrics()
                usd_rate = self._parse_usd_rate()
                self.log(f"💱 Normalisasi ke full amount IDR (USD={usd_rate})", "info")

                if not os.path.exists(csv_s):
                    raise FileNotFoundError(f"CSV sumber tidak ditemukan: {csv_s}")

                try:
                    shutil.copy2(csv_s, csv_raw)
                    self.log(f"✓ Snapshot RAW (sebelum konversi) → {csv_raw}", "dim")
                except Exception as copy_err:
                    raise RuntimeError(f"Gagal membuat snapshot RAW: {copy_err}")

                ok = normalize_to_full_idr(
                    csv_s,
                    selected_metrics=selected,
                    usd_rate=usd_rate,
                    log_callback=slog,
                )

                if ok:
                    try:
                        shutil.copy2(csv_s, csv_full_idr)
                        self.log(f"✓ Simpan hasil FULL IDR → {csv_full_idr}", "dim")
                    except Exception as copy_err:
                        raise RuntimeError(f"Normalisasi berhasil, tetapi gagal simpan CSV FULL IDR: {copy_err}")

                self._set_step(key, "✓ Selesai" if ok else "Gagal",
                               C_SUCCESS if ok else C_DANGER, 1.0)

            elif key == "export":
                self.log("📊 Export 2 CSV + 2 Excel...", "info")

                if not os.path.exists(csv_s):
                    raise FileNotFoundError(f"CSV sumber tidak ditemukan: {csv_s}")

                if not os.path.exists(csv_raw):
                    shutil.copy2(csv_s, csv_raw)
                    self.log("ℹ File RAW belum ada, dibuat dari CSV terkini.", "warn")

                if not os.path.exists(csv_full_idr):
                    shutil.copy2(csv_s, csv_full_idr)
                    self.log("ℹ File FULL IDR belum ada, dibuat dari CSV terkini.", "warn")

                raw_xlsx = os.path.join(ext, "00_SUMMARY_all_companies_raw.xlsx")
                full_idr_xlsx = os.path.join(ext, "00_SUMMARY_all_companies_full_idr.xlsx")

                xlsx_raw = export_to_excel(csv_raw, excel_path=raw_xlsx, log_callback=slog)
                xlsx_full_idr = export_to_excel(csv_full_idr, excel_path=full_idr_xlsx, log_callback=slog)

                if xlsx_raw and xlsx_full_idr:
                    self._set_step(key, "✓ 2 CSV + 2 XLSX", C_SUCCESS, 1.0)
                    self.log(f"✓ RAW CSV         → {csv_raw}", "ok")
                    self.log(f"✓ RAW Excel       → {xlsx_raw}", "ok")
                    self.log(f"✓ FULL IDR CSV    → {csv_full_idr}", "ok")
                    self.log(f"✓ FULL IDR Excel  → {xlsx_full_idr}", "ok")
                else:
                    self._set_step(key, "Gagal", C_DANGER)

        except Exception as e:
            self._set_step(key, "Error!", C_DANGER)
            self.log(f"✗ Error '{key}': {e}", "err")


def run():
    app = IDXSuperApp()
    app.mainloop()


if __name__ == "__main__":
    run()
