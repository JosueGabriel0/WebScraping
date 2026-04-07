"""
SENAMHI Scraper v12 — Edición Profesional
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
UI rediseñada: estética dashboard gubernamental
  · Tipografía clara, espaciado generoso
  · Paleta azul institucional + blanco + gris
  · Panel de control con tarjetas de métricas
  · Log estructurado con timestamps y niveles
  · Barra de progreso por estación y total
  · Sin cambios en lógica de scraping (v11)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import os
import csv
import time
import subprocess

# ── Paleta institucional ───────────────────────────────────────────────────────
C = {
    "navy":     "#1a3a5c",   # azul institucional oscuro
    "blue":     "#1e6bb8",   # azul principal
    "blue_lt":  "#2e86de",   # azul claro / hover
    "sky":      "#e8f1fa",   # fondo azul muy claro
    "white":    "#ffffff",
    "gray_100": "#f4f6f9",   # fondo general
    "gray_200": "#e9ecef",   # bordes, separadores
    "gray_400": "#adb5bd",   # texto secundario
    "gray_700": "#495057",   # texto principal
    "gray_900": "#212529",   # texto oscuro / títulos
    "green":    "#1a7a4a",   # éxito
    "green_lt": "#d4edda",   # fondo éxito
    "amber":    "#b45309",   # advertencia
    "amber_lt": "#fef3c7",   # fondo advertencia
    "red":      "#b91c1c",   # error
    "red_lt":   "#fee2e2",   # fondo error
    "info":     "#1e6bb8",   # info = azul
    "info_lt":  "#dbeafe",   # fondo info
}

DEP_SLUGS = {
    "Amazonas": "amazonas", "Ancash": "ancash", "Apurimac": "apurimac",
    "Arequipa": "arequipa", "Ayacucho": "ayacucho", "Cajamarca": "cajamarca",
    "Cusco": "cusco", "Huancavelica": "huancavelica", "Huánuco": "huanuco",
    "Ica": "ica", "Junín": "junin", "La Libertad": "la-libertad",
    "Lambayeque": "lambayeque", "Lima / Callao": "lima", "Loreto": "loreto",
    "Madre de Dios": "madre-de-dios", "Moquegua": "moquegua", "Pasco": "pasco",
    "Piura": "piura", "Puno": "puno", "San Martin": "san-martin",
    "Tacna": "tacna", "Tumbes": "tumbes", "Ucayali": "ucayali",
}

CHROME_EXE_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
DEBUG_PORT = 9222
DEBUG_DIR  = r"C:\chrome_debug"

SEL_MARCADOR    = "div.leaflet-marker-pane img.leaflet-marker-icon"
SEL_POPUP       = "div.leaflet-popup"
SEL_POPUP_CLOSE = "a.leaflet-popup-close-button"

BASE_DIR = os.path.abspath(os.path.join(os.getcwd(), "senamhi_datos"))


def limpiar_nombre(texto: str, max_len: int = 40) -> str:
    limpio = "".join(c if c.isalnum() or c in " _-" else "_" for c in texto)
    return limpio.strip().replace(" ", "_")[:max_len]


def carpeta_estacion(departamento: str, estacion: str) -> str:
    dep_dir = os.path.join(BASE_DIR, limpiar_nombre(departamento, 30))
    est_dir = os.path.join(dep_dir, limpiar_nombre(estacion, 40))
    os.makedirs(est_dir, exist_ok=True)
    return est_dir


# ── Widget: barra de progreso custom ─────────────────────────────────────────
class ProgressBar(tk.Canvas):
    def __init__(self, parent, width=400, height=10, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=C["gray_200"], highlightthickness=0, **kwargs)
        self._default_w = width
        self._h         = height
        self._val       = 0.0
        # diferir el primer dibujo hasta que el widget esté registrado
        self.after(50, self._draw)

    def set(self, value):
        self._val = max(0.0, min(1.0, value))
        self._draw()

    def _draw(self):
        self.delete("all")
        # usar ancho real si ya está disponible
        w = self.winfo_width()
        if w < 2:
            w = self._default_w
        self.create_rectangle(0, 0, w, self._h, fill=C["gray_200"], outline="")
        fw = int(w * self._val)
        if fw > 0:
            self.create_rectangle(0, 0, fw, self._h, fill=C["blue"], outline="")
        pct = f"{int(self._val * 100)}%"
        self.create_text(w - 28, self._h // 2 + 1,
                         text=pct, font=("Segoe UI", 7, "bold"),
                         fill=C["gray_700"])


# ── Widget: tarjeta métrica ───────────────────────────────────────────────────
class MetricCard(tk.Frame):
    def __init__(self, parent, label, icon, value="0",
                 accent=None, accent_lt=None, **kwargs):
        bg_card = C["white"]
        super().__init__(parent, bg=bg_card,
                         highlightbackground=C["gray_200"],
                         highlightthickness=1, **kwargs)
        self._accent = accent or C["blue"]

        # franja de color superior
        stripe = tk.Frame(self, bg=self._accent, height=4)
        stripe.pack(fill="x")

        inner = tk.Frame(self, bg=bg_card, padx=14, pady=10)
        inner.pack(fill="both", expand=True)

        top_row = tk.Frame(inner, bg=bg_card)
        top_row.pack(fill="x")
        tk.Label(top_row, text=icon, bg=bg_card, fg=self._accent,
                 font=("Segoe UI Emoji", 14)).pack(side="left")
        tk.Label(top_row, text=label.upper(), bg=bg_card,
                 fg=C["gray_400"], font=("Segoe UI", 7, "bold")).pack(
                 side="left", padx=(6, 0), pady=2)

        self._val_lbl = tk.Label(inner, text=value, bg=bg_card,
                                  fg=C["gray_900"], font=("Segoe UI", 26, "bold"))
        self._val_lbl.pack(anchor="w", pady=(2, 0))

    def set(self, val):
        self._val_lbl.config(text=str(val))


# ── Widget: badge de estado ───────────────────────────────────────────────────
class StatusBadge(tk.Frame):
    STATES = {
        "idle":    ("●  Inactivo",  C["gray_400"], C["gray_200"]),
        "running": ("●  Ejecutando", C["blue"],     C["info_lt"]),
        "ok":      ("●  Completado", C["green"],    C["green_lt"]),
        "error":   ("●  Error",      C["red"],      C["red_lt"]),
    }
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C["gray_100"], **kwargs)
        self._lbl = tk.Label(self, font=("Segoe UI", 8, "bold"), padx=10, pady=4)
        self._lbl.pack()
        self.set("idle")

    def set(self, state):
        text, fg, bg = self.STATES.get(state, self.STATES["idle"])
        self._lbl.config(text=text, fg=fg, bg=bg)
        self.config(bg=bg)


# ── Aplicación principal ──────────────────────────────────────────────────────
class SenamhiScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SENAMHI — Sistema de Extracción de Datos Hidrometeorológicos")
        self.root.geometry("980x700")
        self.root.configure(bg=C["gray_100"])
        self.root.resizable(True, True)
        os.makedirs(BASE_DIR, exist_ok=True)

        self._scraping = False
        self._setup_styles()
        self._build_ui()

    # ── estilos ttk ──────────────────────────────────────────────────────────
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Pro.TCombobox",
                    fieldbackground=C["white"],
                    background=C["white"],
                    foreground=C["gray_900"],
                    arrowcolor=C["blue"],
                    bordercolor=C["gray_200"],
                    lightcolor=C["gray_200"],
                    darkcolor=C["gray_200"],
                    selectbackground=C["sky"],
                    selectforeground=C["gray_900"],
                    font=("Segoe UI", 9))
        s.map("Pro.TCombobox",
              fieldbackground=[("readonly", C["white"])],
              foreground=[("readonly", C["gray_900"])])

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):

        # ── HEADER / navbar ───────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=C["navy"], height=56)
        header.pack(fill="x")
        header.pack_propagate(False)

        # logo / título
        title_frame = tk.Frame(header, bg=C["navy"])
        title_frame.pack(side="left", padx=20, pady=10)

        tk.Label(title_frame, text="SENAMHI", bg=C["navy"], fg=C["white"],
                 font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Frame(title_frame, bg=C["blue_lt"], width=2, height=28).pack(
            side="left", padx=12)
        tk.Label(title_frame,
                 text="Sistema de Extracción de Datos Hidrometeorológicos",
                 bg=C["navy"], fg="#a8c4e0",
                 font=("Segoe UI", 9)).pack(side="left")

        # versión
        tk.Label(header, text="v12 PRO", bg=C["blue"], fg=C["white"],
                 font=("Segoe UI", 8, "bold"), padx=10, pady=6).pack(
                 side="right", padx=16, pady=10)

        # ── SUBHEADER: controles ───────────────────────────────────────────────
        ctrl_bar = tk.Frame(self.root, bg=C["white"],
                            highlightbackground=C["gray_200"], highlightthickness=1)
        ctrl_bar.pack(fill="x")

        inner_ctrl = tk.Frame(ctrl_bar, bg=C["white"])
        inner_ctrl.pack(padx=20, pady=10, anchor="w")

        # departamento
        tk.Label(inner_ctrl, text="Departamento:", bg=C["white"],
                 fg=C["gray_700"], font=("Segoe UI", 9, "bold")).pack(side="left")
        self.combo_dep = ttk.Combobox(inner_ctrl, values=list(DEP_SLUGS.keys()),
                                       state="readonly", style="Pro.TCombobox", width=18)
        self.combo_dep.pack(side="left", padx=(6, 20))
        self.combo_dep.current(list(DEP_SLUGS.keys()).index("Puno"))

        # botones de acción
        def _make_btn(parent, text, cmd, primary=False):
            if primary:
                b = tk.Button(parent, text=text, command=cmd,
                              bg=C["blue"], fg=C["white"],
                              activebackground=C["blue_lt"], activeforeground=C["white"],
                              relief="flat", font=("Segoe UI", 9, "bold"),
                              padx=14, pady=6, cursor="hand2")
                b.bind("<Enter>", lambda e: b.config(bg=C["blue_lt"]))
                b.bind("<Leave>", lambda e: b.config(bg=C["blue"]))
            else:
                b = tk.Button(parent, text=text, command=cmd,
                              bg=C["white"], fg=C["gray_700"],
                              activebackground=C["sky"], activeforeground=C["blue"],
                              relief="flat", font=("Segoe UI", 9),
                              padx=14, pady=6, cursor="hand2",
                              highlightbackground=C["gray_200"], highlightthickness=1)
                b.bind("<Enter>", lambda e: b.config(bg=C["sky"], fg=C["blue"]))
                b.bind("<Leave>", lambda e: b.config(bg=C["white"], fg=C["gray_700"]))
            return b

        self.btn_chrome = _make_btn(inner_ctrl, "⚡  Abrir Chrome", self.abrir_chrome)
        self.btn_chrome.pack(side="left", padx=3)

        self.btn_scrape = _make_btn(inner_ctrl, "▶  Iniciar Scraping",
                                    self.start_scrape, primary=True)
        self.btn_scrape.pack(side="left", padx=3)

        self.btn_folder = _make_btn(inner_ctrl, "📂  Ver Carpeta", self.open_folder)
        self.btn_folder.pack(side="left", padx=3)

        # badge de estado (derecha)
        self.badge = StatusBadge(inner_ctrl)
        self.badge.pack(side="left", padx=(20, 0))

        # ── BODY ──────────────────────────────────────────────────────────────
        body = tk.Frame(self.root, bg=C["gray_100"])
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # ── fila de métricas ───────────────────────────────────────────────
        metrics_row = tk.Frame(body, bg=C["gray_100"])
        metrics_row.pack(fill="x", pady=(0, 14))

        self.card_csv = MetricCard(metrics_row, "Archivos CSV",
                                   "📄", "0", C["green"], C["green_lt"])
        self.card_err = MetricCard(metrics_row, "Errores",
                                   "⚠", "0", C["red"], C["red_lt"])
        self.card_est = MetricCard(metrics_row, "Estaciones",
                                   "📍", "0", C["blue"], C["info_lt"])
        self.card_act = MetricCard(metrics_row, "En Proceso",
                                   "⟳", "—", C["amber"], C["amber_lt"])
        for card in (self.card_csv, self.card_err, self.card_est, self.card_act):
            card.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # ── progreso ───────────────────────────────────────────────────────
        prog_frame = tk.Frame(body, bg=C["white"],
                              highlightbackground=C["gray_200"], highlightthickness=1)
        prog_frame.pack(fill="x", pady=(0, 14))

        prog_inner = tk.Frame(prog_frame, bg=C["white"])
        prog_inner.pack(fill="x", padx=18, pady=12)

        # total
        row1 = tk.Frame(prog_inner, bg=C["white"])
        row1.pack(fill="x", pady=(0, 8))
        tk.Label(row1, text="Progreso total", bg=C["white"],
                 fg=C["gray_700"], font=("Segoe UI", 8, "bold")).pack(side="left")
        self.lbl_prog_total = tk.Label(row1, text="0 / 0 estaciones",
                                        bg=C["white"], fg=C["gray_400"],
                                        font=("Segoe UI", 8))
        self.lbl_prog_total.pack(side="right")
        self.prog_total = ProgressBar(prog_inner, height=10)
        self.prog_total.pack(fill="x", pady=(2, 0))

        # estación actual
        row2 = tk.Frame(prog_inner, bg=C["white"])
        row2.pack(fill="x", pady=(10, 0))
        tk.Label(row2, text="Estación actual — meses", bg=C["white"],
                 fg=C["gray_700"], font=("Segoe UI", 8, "bold")).pack(side="left")
        self.lbl_prog_est = tk.Label(row2, text="—", bg=C["white"],
                                      fg=C["gray_400"], font=("Segoe UI", 8))
        self.lbl_prog_est.pack(side="right")
        self.prog_est = ProgressBar(prog_inner, height=6)
        self.prog_est.pack(fill="x", pady=(2, 0))

        # ── LOG ───────────────────────────────────────────────────────────
        log_card = tk.Frame(body, bg=C["white"],
                            highlightbackground=C["gray_200"], highlightthickness=1)
        log_card.pack(fill="both", expand=True)

        log_hdr = tk.Frame(log_card, bg=C["gray_100"],
                           highlightbackground=C["gray_200"], highlightthickness=1)
        log_hdr.pack(fill="x")
        tk.Label(log_hdr, text="Registro de actividad", bg=C["gray_100"],
                 fg=C["gray_900"], font=("Segoe UI", 9, "bold"),
                 padx=14, pady=8).pack(side="left")
        self.lbl_status = tk.Label(log_hdr, text="Sistema listo.",
                                    bg=C["gray_100"], fg=C["gray_400"],
                                    font=("Segoe UI", 8), padx=14)
        self.lbl_status.pack(side="right")

        log_body = tk.Frame(log_card, bg=C["white"])
        log_body.pack(fill="both", expand=True)

        self.log_area = tk.Text(log_body, font=("Consolas", 8),
                                bg=C["white"], fg=C["gray_700"],
                                relief="flat", wrap="word", state="disabled",
                                padx=14, pady=10,
                                selectbackground=C["sky"])
        scroll = tk.Scrollbar(log_body, command=self.log_area.yview,
                               bg=C["gray_100"], troughcolor=C["gray_200"])
        self.log_area.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_area.pack(fill="both", expand=True)

        # tags semánticos
        self.log_area.tag_config("ok",   foreground=C["green"],
                                          font=("Consolas", 8, "bold"))
        self.log_area.tag_config("err",  foreground=C["red"],
                                          font=("Consolas", 8, "bold"))
        self.log_area.tag_config("warn", foreground=C["amber"])
        self.log_area.tag_config("info", foreground=C["blue"])
        self.log_area.tag_config("dim",  foreground=C["gray_400"])
        self.log_area.tag_config("hdr",  foreground=C["navy"],
                                          font=("Consolas", 8, "bold"))
        self.log_area.tag_config("ts",   foreground=C["gray_400"])

        # ── FOOTER ────────────────────────────────────────────────────────
        footer = tk.Frame(self.root, bg=C["navy"], height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(footer, text=f"Datos guardados en: {BASE_DIR}",
                 bg=C["navy"], fg="#a8c4e0",
                 font=("Segoe UI", 7)).pack(side="left", padx=16, pady=6)
        tk.Label(footer, text="Servicio Nacional de Meteorología e Hidrología del Perú",
                 bg=C["navy"], fg="#a8c4e0",
                 font=("Segoe UI", 7)).pack(side="right", padx=16, pady=6)

    # ── logging ──────────────────────────────────────────────────────────────
    def log(self, msg, tag=None):
        if tag is None:
            if any(x in msg for x in ["✅", "💾", "✔"]):
                tag = "ok"
            elif any(x in msg for x in ["❌", "Error", "ERROR"]):
                tag = "err"
            elif any(x in msg for x in ["⚠", "sin ", "Sin "]):
                tag = "warn"
            elif msg.startswith("[") or any(x in msg for x in ["🖱", "📅"]):
                tag = "info"
            elif msg.startswith("  →") or msg.startswith("   "):
                tag = "dim"
        self.root.after(0, lambda: self._write(msg, tag))
        print(msg)

    def _write(self, msg, tag=None):
        self.log_area.config(state="normal")
        ts = time.strftime("%H:%M:%S")
        self.log_area.insert("end", f"{ts}  ", "ts")
        self.log_area.insert("end", msg + "\n", tag or "")
        self.log_area.see("end")
        self.log_area.config(state="disabled")

    def status(self, msg, color="orange"):
        col_map = {"orange": C["amber"], "green": C["green"],
                   "red": C["red"], "blue": C["blue"]}
        c = col_map.get(color, color)
        self.root.after(0, lambda: self.lbl_status.config(text=msg, foreground=c))

    def _update_stats(self, csv=None, err=None, est=None, est_nombre=None,
                       prog_total=None, prog_est=None,
                       total_est=None, actual_est=None, total_mes=None, actual_mes=None):
        def _do():
            if csv is not None:
                self.card_csv.set(csv)
            if err is not None:
                self.card_err.set(err)
            if est is not None:
                self.card_est.set(est)
            if est_nombre is not None:
                self.card_act.set(est_nombre)
            if prog_total is not None:
                self.prog_total.set(prog_total)
            if prog_est is not None:
                self.prog_est.set(prog_est)
            if total_est is not None and actual_est is not None:
                self.lbl_prog_total.config(
                    text=f"{actual_est} / {total_est} estaciones")
            if total_mes is not None and actual_mes is not None:
                self.lbl_prog_est.config(
                    text=f"{actual_mes} / {total_mes} meses")
        self.root.after(0, _do)

    def open_folder(self):
        try: os.startfile(BASE_DIR)
        except: pass

    def _disable(self):
        self.btn_scrape.config(state="disabled")
        self.badge.set("running")
        self._scraping = True

    def _enable(self):
        def _do():
            self.btn_scrape.config(state="normal")
            self.badge.set("ok")
            self._scraping = False
        self.root.after(0, _do)

    # ── Selenium helpers ──────────────────────────────────────────────────────
    def _connect_driver(self):
        opts = Options()
        opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
        return webdriver.Chrome(options=opts)

    def _set_download_dir(self, driver, path: str):
        try:
            driver.execute_cdp_cmd("Page.setDownloadBehavior",
                                   {"behavior": "allow", "downloadPath": path})
        except Exception as e:
            self.log(f"  ⚠ CDP download: {e}")

    def _entrar_iframe_mapa(self, driver, dep_slug):
        driver.switch_to.default_content()
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            src = iframe.get_attribute("src") or ""
            if "mapa-estaciones" in src and dep_slug in src:
                driver.switch_to.frame(iframe)
                return True
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            src = iframe.get_attribute("src") or ""
            if "mapa-estaciones" in src:
                driver.switch_to.frame(iframe)
                return True
        return False

    def _cerrar_popup(self, driver):
        try:
            btn = driver.find_element(By.CSS_SELECTOR, SEL_POPUP_CLOSE)
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.8)
        except: pass

    def _esperar_csv(self, directorio: str, antes: set, timeout=15):
        t0 = time.time()
        while time.time() - t0 < timeout:
            ahora = {f for f in os.listdir(directorio)
                     if f.endswith(".csv") and not f.endswith(".crdownload")}
            nuevos = ahora - antes
            if nuevos:
                return nuevos.pop()
            time.sleep(0.5)
        return None

    def _renombrar(self, directorio: str, archivo: str, mes: str):
        destino = os.path.join(directorio, f"{mes}.csv")
        origen  = os.path.join(directorio, archivo)
        if os.path.exists(destino):
            destino = os.path.join(directorio, f"{mes}_dup.csv")
        try: os.rename(origen, destino)
        except: pass

    def _guardar_tabla_como_csv(self, driver, directorio: str, mes: str) -> bool:
        try:
            iframes_reales = driver.find_elements(By.TAG_NAME, "iframe")
            filas = None
            entro_iframe_interno = False

            if iframes_reales:
                try:
                    driver.switch_to.frame(iframes_reales[0])
                    entro_iframe_interno = True
                    time.sleep(1.5)

                    body_txt = driver.execute_script(
                        "return document.body ? document.body.innerText : '';")
                    if any(e in body_txt for e in
                           ['Fatal error', 'TypeError', 'Notice:', 'Warning:']):
                        self.log(f"    ⚠ {mes}: error PHP en servidor")
                        driver.switch_to.parent_frame()
                        nombre_err = os.path.join(directorio, f"{mes}_ERROR.csv")
                        with open(nombre_err, "w", encoding="utf-8-sig") as f:
                            f.write("mes,estado\n" + mes + ",ERROR_PHP\n")
                        return False

                    filas = driver.execute_script("""
                        var tablas = Array.from(document.querySelectorAll('table'));
                        if (!tablas.length) return null;
                        tablas.sort(function(a,b){
                            return b.querySelectorAll('tr').length
                                 - a.querySelectorAll('tr').length;
                        });
                        var tabla = tablas[0];
                        var rows = [];
                        var trs = tabla.querySelectorAll('tr');
                        for (var j = 0; j < trs.length; j++) {
                            var celdas = Array.from(trs[j].querySelectorAll('th, td'))
                                .map(function(c){ return c.innerText.trim(); });
                            if (celdas.some(function(c){ return c !== ''; }))
                                rows.push(celdas);
                        }
                        return rows.length >= 1 ? rows : null;
                    """)
                    driver.switch_to.parent_frame()
                except Exception as e_inner:
                    self.log(f"    ⚠ Error iframe interno: {e_inner}")
                    try: driver.switch_to.parent_frame()
                    except: pass
                    filas = None

            if not filas:
                filas = driver.execute_script("""
                    var tablas = Array.from(document.querySelectorAll('table'));
                    if (!tablas.length) return null;
                    tablas.sort(function(a,b){
                        return b.querySelectorAll('tr').length
                             - a.querySelectorAll('tr').length;
                    });
                    var tabla = tablas[0];
                    var rows = [];
                    var trs = tabla.querySelectorAll('tr');
                    for (var j = 0; j < trs.length; j++) {
                        var celdas = Array.from(trs[j].querySelectorAll('th, td'))
                            .map(function(c){ return c.innerText.trim(); });
                        if (celdas.some(function(c){ return c !== ''; }))
                            rows.push(celdas);
                    }
                    return rows.length >= 1 ? rows : null;
                """)

            if not filas:
                self.log(f"    ⚠ {mes}: sin datos en ningún frame")
                return False

            if len(filas) == 1:
                self.log(f"    ⚠ {mes}: solo cabecera (sin datos ese mes)")
                filas.append(["SIN_DATOS"] * len(filas[0]))

            nombre_csv = os.path.join(directorio, f"{mes}.csv")
            with open(nombre_csv, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerows(filas)

            origen = "iframe_interno" if entro_iframe_interno else "frame_popup"
            self.log(f"    💾 {mes}.csv ({len(filas)-1} filas, "
                     f"{len(filas[0])} cols) [{origen}]")
            return True

        except Exception as e:
            self.log(f"    ❌ Error guardando tabla {mes}: {e}")
            return False

    def _click_boton_csv(self, driver) -> bool:
        resultado = driver.execute_script("""
            var textos = ['exportar a csv','exportar csv','export csv','csv'];
            var elementos = Array.from(document.querySelectorAll(
                'button,a,input[type=button],input[type=submit]'));
            for (var i=0;i<elementos.length;i++){
                var el=elementos[i];
                var t=(el.innerText||el.value||el.title||'').toLowerCase().trim();
                for(var j=0;j<textos.length;j++)
                    if(t.indexOf(textos[j])!==-1){el.click();return 'OK:'+t;}
            }
            var todos=Array.from(document.querySelectorAll('*'));
            for(var k=0;k<todos.length;k++){
                var el2=todos[k];
                var attrs=Array.from(el2.attributes)
                    .map(function(a){return a.value.toLowerCase();}).join(' ');
                var txt2=(el2.innerText||'').toLowerCase();
                if((attrs.indexOf('csv')!==-1||txt2.indexOf('exportar')!==-1)&&
                   ['BUTTON','A','INPUT'].indexOf(el2.tagName)!==-1){
                    el2.click();return 'FALLBACK:'+el2.tagName;
                }
            }
            return null;
        """)
        if resultado:
            self.log(f"    🖱 Botón CSV: {resultado}")
            return True
        for xp in [
            "//*[contains(translate(normalize-space(.),'EXPORTAR A CSV','exportar a csv'),'exportar a csv')]",
            "//*[contains(translate(normalize-space(.),'EXPORTAR','exportar'),'exportar')]",
        ]:
            try:
                for el in driver.find_elements(By.XPATH, xp):
                    txt = (el.text or el.get_attribute("value") or "").lower()
                    if "csv" in txt or "exportar" in txt:
                        driver.execute_script("arguments[0].click();", el)
                        self.log(f"    🖱 Botón vía XPath: '{el.text}'")
                        return True
            except: continue
        return False

    # ── Abrir Chrome ──────────────────────────────────────────────────────────
    def abrir_chrome(self):
        dep      = self.combo_dep.get()
        dep_slug = DEP_SLUGS[dep]
        url      = f"https://www.senamhi.gob.pe/?p=estaciones&dp={dep_slug}"
        os.makedirs(DEBUG_DIR, exist_ok=True)
        for exe in CHROME_EXE_PATHS:
            if os.path.exists(exe):
                subprocess.Popen([exe,
                    f"--remote-debugging-port={DEBUG_PORT}",
                    f"--user-data-dir={DEBUG_DIR}", url])
                self.log(f"✅ Chrome abierto → {url}", "ok")
                self.status("Chrome abierto. Espera que los pines aparezcan.", "blue")
                self.badge.set("running")
                return
        self.log("❌ Chrome no encontrado en rutas predefinidas.", "err")
        self.badge.set("error")

    # ── Scraping ──────────────────────────────────────────────────────────────
    def start_scrape(self):
        self._disable()
        dep = self.combo_dep.get()
        self.status(f"Iniciando extracción: {dep}...", "orange")
        threading.Thread(target=self._run_scrape, args=(dep,), daemon=True).start()

    def _run_scrape(self, departamento):
        driver = None
        try:
            dep_slug = DEP_SLUGS[departamento]
            self.log(f"Conectando al navegador Chrome...")
            driver = self._connect_driver()
            self.log(f"Página activa: {driver.title}", "info")

            if not self._entrar_iframe_mapa(driver, dep_slug):
                self.status("No se encontró el iframe del mapa.", "red")
                self.badge.set("error")
                return

            self.log("Esperando marcadores de estaciones en el mapa...")
            WebDriverWait(driver, 20).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, SEL_MARCADOR)) > 0
            )

            marcadores = driver.find_elements(By.CSS_SELECTOR, SEL_MARCADOR)
            total      = len(marcadores)
            self._update_stats(est=total, total_est=total, actual_est=0)
            self.log(f"✅ {total} estaciones encontradas en {departamento}.", "ok")

            total_csv = 0
            total_err = 0

            for i in range(total):
                try:
                    marcadores = driver.find_elements(By.CSS_SELECTOR, SEL_MARCADOR)
                    if i >= len(marcadores):
                        continue

                    nombre = marcadores[i].get_attribute("title") or f"Estacion_{i+1}"
                    nombre_corto = nombre[:22] + "…" if len(nombre) > 22 else nombre
                    self._update_stats(
                        csv=total_csv, err=total_err,
                        est_nombre=nombre_corto,
                        prog_total=i / total if total else 0,
                        total_est=total, actual_est=i + 1
                    )
                    self.status(f"Procesando [{i+1}/{total}]: {nombre[:40]}", "orange")
                    self.log(f"\n── [{i+1}/{total}] {nombre}", "hdr")

                    est_dir = carpeta_estacion(departamento, nombre)
                    self._set_download_dir(driver, est_dir)

                    driver.execute_script("arguments[0].click();", marcadores[i])
                    time.sleep(3)

                    try:
                        WebDriverWait(driver, 8).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, SEL_POPUP))
                        )
                    except:
                        self.log(f"  ⚠ Sin popup para esta estación")
                        total_err += 1
                        continue

                    popup_iframes = driver.find_elements(
                        By.CSS_SELECTOR, f"{SEL_POPUP} iframe")
                    if popup_iframes:
                        driver.switch_to.frame(popup_iframes[0])
                        time.sleep(2)
                    else:
                        self.log(f"  ⚠ Sin iframe en popup — contenido directo")

                    botones_info = driver.execute_script("""
                        return Array.from(document.querySelectorAll(
                            'button,a,input[type=button],input[type=submit]'))
                            .map(function(e){
                                return (e.innerText||e.value||e.title||'').trim();
                            }).filter(function(t){ return t.length > 0; });
                    """)
                    self.log(f"  Botones: {botones_info}", "dim")

                    tab_ok = False
                    for xp in [
                        "//a[contains(translate(normalize-space(.),'TABLA','tabla'),'tabla')]",
                        "//li/a[contains(translate(normalize-space(.),'TABLA','tabla'),'tabla')]",
                    ]:
                        elems = driver.find_elements(By.XPATH, xp)
                        if elems:
                            driver.execute_script("arguments[0].click();", elems[0])
                            tab_ok = True
                            time.sleep(2)
                            break

                    if not tab_ok:
                        self.log(f"  ⚠ Sin pestaña Tabla")
                        total_err += 1
                        driver.switch_to.default_content()
                        self._entrar_iframe_mapa(driver, dep_slug)
                        self._cerrar_popup(driver)
                        continue

                    opciones = driver.execute_script("""
                        var sel = document.querySelector('select');
                        if (!sel) return [];
                        return Array.from(sel.options)
                            .filter(function(o){ return o.value; })
                            .map(function(o){ return [o.value, o.text.trim()]; });
                    """)
                    if not opciones:
                        self.log(f"  ⚠ Sin opciones en el selector de meses")
                        driver.switch_to.default_content()
                        self._entrar_iframe_mapa(driver, dep_slug)
                        self._cerrar_popup(driver)
                        continue
                    self.log(f"  📅 {len(opciones)} meses disponibles: "
                             f"{opciones[0][0]} → {opciones[-1][0]}", "info")

                    for idx_mes, (valor, texto) in enumerate(opciones):
                        self._update_stats(
                            prog_est=(idx_mes / len(opciones)) if opciones else 0,
                            total_mes=len(opciones), actual_mes=idx_mes + 1
                        )
                        try:
                            csvs_antes = {f for f in os.listdir(est_dir)
                                          if f.endswith(".csv")
                                          and not f.endswith(".crdownload")}

                            driver.execute_script("""
                                var sel = document.querySelector('select');
                                if (sel) {
                                    sel.value = arguments[0];
                                    sel.dispatchEvent(new Event('change',{bubbles:true}));
                                    sel.dispatchEvent(new Event('input', {bubbles:true}));
                                }
                            """, valor)
                            time.sleep(3.5)

                            btn_ok = self._click_boton_csv(driver)

                            if btn_ok:
                                time.sleep(4)
                                archivo = self._esperar_csv(est_dir, csvs_antes, timeout=8)
                                if archivo:
                                    self._renombrar(est_dir, archivo, valor)
                                    self.log(f"    ✅ {valor}.csv (descarga directa)", "ok")
                                    total_csv += 1
                                else:
                                    self.log(f"    ⚠ {valor} → sin descarga, extrayendo tabla...")
                                    ok = self._guardar_tabla_como_csv(driver, est_dir, valor)
                                    if ok: total_csv += 1
                                    else:  total_err += 1
                            else:
                                self.log(f"    ⚠ {valor} → sin botón, extrayendo tabla...")
                                ok = self._guardar_tabla_como_csv(driver, est_dir, valor)
                                if ok: total_csv += 1
                                else:  total_err += 1

                            self._update_stats(csv=total_csv, err=total_err)
                            time.sleep(0.5)

                        except Exception as e:
                            self.log(f"    ❌ {valor}: {e}", "err")
                            total_err += 1

                    self._update_stats(prog_est=1.0)
                    self.log(f"  → Estación completada. Total acumulado: "
                             f"{total_csv} archivos", "dim")

                    driver.switch_to.default_content()
                    self._entrar_iframe_mapa(driver, dep_slug)
                    self._cerrar_popup(driver)
                    time.sleep(1)

                except Exception as e:
                    self.log(f"  [{i+1}] Error inesperado: {e}", "err")
                    total_err += 1
                    try:
                        driver.switch_to.default_content()
                        self._entrar_iframe_mapa(driver, dep_slug)
                    except: pass
                    self._cerrar_popup(driver)

            self._update_stats(csv=total_csv, err=total_err,
                               prog_total=1.0, prog_est=1.0,
                               est_nombre="Completado")
            msg = (f"Extracción finalizada: {total_csv} archivos guardados, "
                   f"{total_err} errores — {total} estaciones procesadas.")
            self.log(f"\n✅ {msg}", "ok")
            self.log(f"   Carpeta: {BASE_DIR}", "dim")
            self.status(msg, "green")

        except Exception as e:
            self.log(f"❌ Error crítico: {e}", "err")
            import traceback
            self.log(traceback.format_exc(), "err")
            self.status("Error crítico durante la extracción.", "red")
            self.root.after(0, lambda: self.badge.set("error"))
        finally:
            try: driver.switch_to.default_content()
            except: pass
            self._enable()


if __name__ == "__main__":
    root = tk.Tk()
    app  = SenamhiScraperApp(root)
    root.mainloop()