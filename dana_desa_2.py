#!/usr/bin/env python3
"""Buku Kas Dana Desa - aplikasi desktop offline, 1 file, tanpa library tambahan.
Jalankan: python dana_desa.py   (data tersimpan di dana_desa_data.json di samping file ini)
"""
import json, os, sys, zipfile, datetime, random, string
from xml.sax.saxutils import escape

BASE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
DATAFILE = os.path.join(BASE, "dana_desa_data.json")
KAT = {"Pemasukan": ["Dana Desa (APBN)", "Alokasi Dana Desa (ADD)", "Pendapatan Asli Desa", "Bantuan Provinsi/Kabupaten", "Lainnya"],
       "Pengeluaran": ["Alat Tulis Kantor", "Pembangunan/Perbaikan", "Honor/Operasional",
                       "Kesehatan/Posyandu", "Kegiatan Warga", "Lainnya"]}


def rp(n):
    return ("-" if n < 0 else "") + "Rp " + f"{abs(int(n)):,}".replace(",", ".")


def gen_nb(jenis, data):
    """Nomor bukti acak, contoh: KLR-20261006-7QX4 (dijamin belum dipakai)."""
    used = {t.get("nb") for t in data}
    pre = "MSK" if jenis == "Pemasukan" else "KLR"
    while True:
        n = f"{pre}-{datetime.date.today():%Y%m%d}-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=4))
        if n not in used:
            return n


def load():
    try:
        with open(DATAFILE, encoding="utf-8") as f:
            d = json.load(f)
        return d.get("data", []), d.get("keg", [])
    except Exception:
        return [], []


def save(data, keg):
    tmp = DATAFILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"v": 1, "data": data, "keg": keg}, f, ensure_ascii=False, indent=1)
    os.replace(tmp, DATAFILE)


# ---------- penulis xlsx minimal (tanpa openpyxl) ----------
def _col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _sheet(rows, widths):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    out.append("<cols>" + "".join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths)) + "</cols><sheetData>")
    for r, row in enumerate(rows, 1):
        out.append(f'<row r="{r}">')
        for c, v in enumerate(row):
            ref = f"{_col(c)}{r}"
            if v is None or v == "":
                continue
            if isinstance(v, (int, float)):
                out.append(f'<c r="{ref}"><v>{v}</v></c>')
            else:
                out.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>')
        out.append("</row>")
    out.append("</sheetData></worksheet>")
    return "".join(out)


def write_xlsx(path, sheets):
    """sheets = [(nama, rows, lebar_kolom)]"""
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(len(sheets)))
          + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          + "".join(f'<sheet name="{escape(n)}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, (n, _, _) in enumerate(sheets))
          + "</sheets></workbook>")
    wbr = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           + "".join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>' for i in range(len(sheets)))
           + "</Relationships>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbr)
        for i, (_, rows, w) in enumerate(sheets):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", _sheet(rows, w))


def tgl_buat(k):
    if k.get("buat"):
        return k["buat"]
    try:
        return datetime.date.fromtimestamp(k["id"] / 1000).isoformat()
    except Exception:
        return ""


def keg_status(k, data):
    """(selesai?, tanggal_selesai). Selesai = ditandai manual, atau terpakai >= anggaran."""
    if k.get("selesai"):
        return True, k["selesai"]
    if k["anggaran"] > 0 and real(data, k["nama"]) >= k["anggaran"]:
        cum = 0
        for t in sorted((t for t in data if t["jenis"] == "Pengeluaran" and t.get("keg") == k["nama"]), key=lambda t: (t["tgl"], t["id"])):
            cum += t["jml"]
            if cum >= k["anggaran"]:
                return True, t["tgl"]
    return False, ""


def build_sheets(view, keg, alldata):
    s = sorted(view, key=lambda t: (t["tgl"], t["id"]))
    rows = [["No", "Tanggal", "No. Bukti", "Jenis", "Kategori", "Kegiatan", "Keterangan", "Pemasukan (Rp)", "Pengeluaran (Rp)", "Saldo (Rp)"]]
    saldo = ti = to = 0
    for i, t in enumerate(s, 1):
        inn = t["jenis"] == "Pemasukan"
        saldo += t["jml"] if inn else -t["jml"]
        ti += t["jml"] if inn else 0
        to += 0 if inn else t["jml"]
        rows.append([i, t["tgl"], t.get("nb", ""), t["jenis"], t["kat"], t.get("keg", ""), t["ket"],
                     t["jml"] if inn else "", "" if inn else t["jml"], saldo])
    rows += [[], ["", "", "", "", "", "", "TOTAL", ti, to, ti - to]]
    g = {}
    for t in s:
        g[(t["jenis"], t["kat"])] = g.get((t["jenis"], t["kat"]), 0) + t["jml"]
    rk = [["Jenis", "Kategori", "Total (Rp)"]] + [[j, c, v] for (j, c), v in sorted(g.items())]
    hd = ["Kegiatan", "Tanggal Buat", "Tanggal Selesai", "Anggaran (Rp)", "Terpakai (Rp)", "Sisa (Rp)", "Terpakai (%)"]
    wait, done = [hd], [hd]
    for k in keg:
        r = real(alldata, k["nama"])
        ok, tsel = keg_status(k, alldata)
        (done if ok else wait).append([k["nama"], tgl_buat(k), tsel, k["anggaran"], r, k["anggaran"] - r,
                                        round(r / k["anggaran"] * 100) if k["anggaran"] else 0])
    w = [30, 14, 16, 16, 16, 16, 14]
    return [("Transaksi", rows, [5, 12, 14, 13, 24, 24, 40, 16, 16, 16]),
            ("Ringkasan Kategori", rk, [14, 28, 16]),
            ("Kegiatan Belum Selesai", wait, w),
            ("Kegiatan Selesai", done, w)]


def real(data, nama):
    return sum(t["jml"] for t in data if t["jenis"] == "Pengeluaran" and t.get("keg") == nama)


# ---------- antarmuka ----------
BG, CARD, LINE = "#eef2f7", "#ffffff", "#dbe3ee"
TEAL, GREEN, RED, BLUE, AMBER, SLATE, INK = "#0f766e", "#16a34a", "#e11d48", "#2563eb", "#f59e0b", "#64748b", "#0f172a"


def main():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    data, keg = load()
    root = tk.Tk()
    root.title("Buku Kas Dana Desa")
    root.geometry("1220x760")
    root.configure(bg=BG)
    F = "Segoe UI"

    st = ttk.Style()
    st.theme_use("clam")
    st.configure(".", font=(F, 10))
    st.configure("TFrame", background=BG)
    st.configure("Card.TFrame", background=CARD)
    st.configure("TLabel", background=BG, foreground=INK)
    st.configure("Card.TLabel", background=CARD, foreground=SLATE, font=(F, 9))
    st.configure("TNotebook", background=BG, borderwidth=0)
    st.configure("TNotebook.Tab", padding=(22, 9), font=(F, 10, "bold"), background="#cbd5e1", foreground=INK)
    st.map("TNotebook.Tab", background=[("selected", TEAL)], foreground=[("selected", "white")])
    st.configure("TLabelframe", background=CARD, bordercolor=LINE, relief="solid")
    st.configure("TLabelframe.Label", background=CARD, foreground=INK, font=(F, 10, "bold"))
    st.configure("TEntry", padding=5, fieldbackground="white")
    st.configure("TCombobox", padding=5, fieldbackground="white")
    st.configure("TRadiobutton", background=CARD, font=(F, 10, "bold"))
    for nm, col in (("In", GREEN), ("Out", RED), ("Wait", AMBER), ("Done", BLUE)):
        st.configure(f"{nm}.Treeview", rowheight=27, background="white", fieldbackground="white", bordercolor=LINE)
        st.configure(f"{nm}.Treeview.Heading", background=col, foreground="white", font=(F, 10, "bold"), padding=6, relief="flat")
        st.map(f"{nm}.Treeview.Heading", background=[("active", col)])
        st.map(f"{nm}.Treeview", background=[("selected", col)], foreground=[("selected", "white")])

    def btn(parent, text, cmd, color, **kw):
        dark = {GREEN: "#15803d", RED: "#be123c", BLUE: "#1d4ed8", TEAL: "#115e59", SLATE: "#475569", AMBER: "#d97706"}[color]
        b = tk.Button(parent, text=text, command=cmd, bg=color, fg="white", activebackground=dark, activeforeground="white",
                      relief="flat", bd=0, padx=14, pady=7, font=(F, 10, "bold"), cursor="hand2", **kw)
        b.bind("<Enter>", lambda e: b.config(bg=dark))
        b.bind("<Leave>", lambda e: b.config(bg=color))
        return b

    # header
    hd = tk.Frame(root, bg=TEAL)
    hd.pack(fill="x")
    tk.Label(hd, text="📒  Buku Kas Dana Desa", bg=TEAL, fg="white", font=(F, 18, "bold")).pack(side="left", padx=20, pady=(12, 2))
    tk.Label(hd, text="Pencatatan pemasukan, pengeluaran & anggaran kegiatan", bg=TEAL, fg="#99f6e4", font=(F, 10)).pack(side="left", padx=4, pady=(18, 0))

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=12, pady=10)
    tab1, tab2 = ttk.Frame(nb), ttk.Frame(nb)
    nb.add(tab1, text="  💰 Transaksi  ")
    nb.add(tab2, text="  🎯 Anggaran Kegiatan  ")

    # kartu ringkasan
    cards = tk.Frame(tab1, bg=BG)
    cards.pack(fill="x", pady=(8, 8))
    val = {}
    for i, (k, t, col) in enumerate((("in", "TOTAL PEMASUKAN", GREEN), ("out", "TOTAL PENGELUARAN", RED), ("saldo", "SALDO", BLUE))):
        c = tk.Frame(cards, bg=col)
        c.grid(row=0, column=i, sticky="we", padx=(0 if i == 0 else 8, 0))
        cards.columnconfigure(i, weight=1)
        tk.Label(c, text=t, bg=col, fg="white", font=(F, 9, "bold")).pack(anchor="w", padx=16, pady=(10, 0))
        val[k] = tk.Label(c, text="", bg=col, fg="white", font=(F, 20, "bold"))
        val[k].pack(anchor="w", padx=16, pady=(0, 10))

    body = tk.Frame(tab1, bg=BG)
    body.pack(fill="both", expand=True)

    # ---- form kiri ----
    left = tk.Frame(body, bg=CARD, highlightbackground=LINE, highlightthickness=1)
    left.pack(side="left", fill="y", padx=(0, 10))
    tk.Label(left, text="Catat Transaksi", bg=CARD, fg=INK, font=(F, 12, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
    jenis = tk.StringVar(value="Pemasukan")
    tgl, nobukti, kat, kg, ket, jml = (tk.StringVar() for _ in range(6))
    tgl.set(datetime.date.today().isoformat())
    fr = tk.Frame(left, bg=CARD)
    fr.pack(anchor="w", padx=12)
    for v in ("Pemasukan", "Pengeluaran"):
        ttk.Radiobutton(fr, text=v, value=v, variable=jenis, command=lambda: set_kat()).pack(side="left", padx=4)

    def field(text, var, widget=ttk.Entry, w=30):
        ttk.Label(left, text=text, style="Card.TLabel").pack(anchor="w", padx=14, pady=(8, 0))
        e = widget(left, textvariable=var, width=w)
        e.pack(fill="x", padx=14)
        return e
    field("Tanggal (YYYY-MM-DD)", tgl)
    ttk.Label(left, text="No. Bukti / Nota (otomatis)", style="Card.TLabel").pack(anchor="w", padx=14, pady=(8, 0))
    nbf = tk.Frame(left, bg=CARD)
    nbf.pack(fill="x", padx=14)
    ttk.Entry(nbf, textvariable=nobukti, state="readonly").pack(side="left", fill="x", expand=True)
    cb_kat = field("Kategori", kat, ttk.Combobox)
    cb_kg = field("Kegiatan (opsional, belum selesai)", kg, ttk.Combobox)
    field("Jumlah (Rp)", jml)
    field("Keterangan", ket)

    def newnb():
        nobukti.set(gen_nb(jenis.get(), data))
    btn(nbf, "🔄", newnb, SLATE).pack(side="left", padx=(6, 0))

    def set_kat():
        cb_kat["values"] = KAT[jenis.get()]
        kat.set("")
        cb_kg.config(state="readonly" if jenis.get() == "Pengeluaran" else "disabled")
        if jenis.get() != "Pengeluaran":
            kg.set("")
        newnb()

    def tambah():
        try:
            d = datetime.date.fromisoformat(tgl.get().strip())
            j = int("".join(c for c in jml.get() if c.isdigit()))
            assert j > 0
        except Exception:
            messagebox.showwarning("Periksa isian", "Tanggal harus format YYYY-MM-DD dan jumlah harus angka.")
            return
        if not kat.get().strip() or not ket.get().strip():
            messagebox.showwarning("Periksa isian", "Kategori dan keterangan wajib diisi.")
            return
        data.append({"id": int(datetime.datetime.now().timestamp() * 1000), "tgl": d.isoformat(), "nb": nobukti.get(),
                     "jenis": jenis.get(), "kat": kat.get().strip(), "keg": kg.get().strip(), "ket": ket.get().strip(), "jml": j})
        save(data, keg)
        for v in (kat, ket, jml):
            v.set("")
        newnb()
        refresh()
    btn(left, "+  Simpan Transaksi", tambah, TEAL).pack(fill="x", padx=14, pady=(16, 14))

    # ---- kanan: filter + 2 tabel ----
    right = tk.Frame(body, bg=BG)
    right.pack(side="left", fill="both", expand=True)
    bln, fj, fk, fg = (tk.StringVar(value="Semua") for _ in range(4))
    flt = tk.Frame(right, bg=CARD, highlightbackground=LINE, highlightthickness=1)
    flt.pack(fill="x", pady=(0, 8))
    cbs = {}
    for i, (t, v, w) in enumerate((("Bulan", bln, 9), ("Jenis", fj, 11), ("Kategori", fk, 20), ("Kegiatan", fg, 24))):
        ttk.Label(flt, text="🔎 " + t if i == 0 else t, style="Card.TLabel").grid(row=0, column=i, sticky="w", padx=8, pady=(6, 0))
        cbs[t] = ttk.Combobox(flt, textvariable=v, state="readonly", width=w)
        cbs[t].grid(row=1, column=i, padx=8, pady=(0, 8))
    cbs["Jenis"]["values"] = ["Semua", "Pemasukan", "Pengeluaran"]

    def reset():
        for v in (bln, fj, fk, fg):
            v.set("Semua")
        refresh()
    btn(flt, "Reset", reset, SLATE).grid(row=1, column=4, padx=8, pady=(0, 8))

    cols = ("tgl", "nb", "kat", "keg", "ket", "jml")
    heads = (("Tanggal", 85), ("No. Bukti", 125), ("Kategori", 130), ("Kegiatan", 140), ("Keterangan", 200), ("Jumlah", 110))

    def mk(parent, title, style, columns, hds, tags, height=5):
        lf = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        lf.pack(fill="both", expand=True, pady=(0, 8))
        h = tk.Label(lf, text=title, bg=CARD, fg=INK, font=(F, 11, "bold"))
        h.pack(anchor="w", padx=10, pady=(8, 4))
        wrap = tk.Frame(lf, bg=CARD)
        wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        t = ttk.Treeview(wrap, columns=columns, show="headings", height=height, style=style)
        for c, (hh, w) in zip(columns, hds):
            t.heading(c, text=hh)
            t.column(c, width=w, anchor="e" if c in ("jml", "ang", "pakai", "sisa", "pct") else "w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(fill="both", expand=True)
        t.tag_configure("odd", background="#f8fafc")
        for tg, col in tags.items():
            t.tag_configure(tg, foreground=col)
        return h, t
    lf_in, tv_in = mk(right, "", "In.Treeview", cols, heads, {"r": GREEN})
    lf_out, tv_out = mk(right, "", "Out.Treeview", cols, heads, {"r": RED})

    def view():
        v = data
        if bln.get() != "Semua":
            v = [t for t in v if t["tgl"].startswith(bln.get())]
        if fj.get() != "Semua":
            v = [t for t in v if t["jenis"] == fj.get()]
        if fk.get() != "Semua":
            v = [t for t in v if t["kat"] == fk.get()]
        if fg.get() != "Semua":
            v = [t for t in v if t.get("keg") == fg.get()]
        return v

    def hapus():
        sel = list(tv_in.selection()) + list(tv_out.selection())
        if sel and messagebox.askyesno("Hapus", f"Hapus {len(sel)} transaksi terpilih?"):
            ids = {int(i) for i in sel}
            data[:] = [t for t in data if t["id"] not in ids]
            save(data, keg)
            refresh()

    def excel():
        v = view()
        if not v:
            messagebox.showinfo("Excel", "Tidak ada data pada filter ini.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                         initialfile=f"Dana-Desa-{bln.get() if bln.get() != 'Semua' else 'semua'}.xlsx")
        if p:
            try:
                write_xlsx(p, build_sheets(v, keg, data))
                messagebox.showinfo("Excel", "File Excel tersimpan:\n" + p)
            except Exception as e:
                messagebox.showerror("Gagal", str(e))

    def cadang():
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")],
                                         initialfile=f"Cadangan-Dana-Desa-{datetime.date.today()}.json")
        if p:
            with open(p, "w", encoding="utf-8") as fh:
                json.dump({"v": 1, "data": data, "keg": keg}, fh, ensure_ascii=False, indent=1)

    def pulihkan():
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            with open(p, encoding="utf-8") as fh:
                o = json.load(fh)
            assert isinstance(o["data"], list) and isinstance(o["keg"], list)
        except Exception:
            messagebox.showerror("Gagal", "File cadangan tidak valid.")
            return
        if messagebox.askyesno("Pulihkan", f"Ganti semua data dengan cadangan ({len(o['data'])} transaksi)?"):
            data[:] = o["data"]
            keg[:] = o["keg"]
            save(data, keg)
            refresh()

    bar = tk.Frame(right, bg=BG)
    bar.pack(fill="x")
    for t, c, col in (("⬇ Unduh Excel", excel, GREEN), ("Hapus Terpilih", hapus, RED), ("Cadangkan", cadang, BLUE), ("Pulihkan", pulihkan, SLATE)):
        btn(bar, t, c, col).pack(side="left", padx=(0, 8))

    # ================= tab kegiatan =================
    g = tk.Frame(tab2, bg=CARD, highlightbackground=LINE, highlightthickness=1)
    g.pack(fill="x", pady=(8, 8))
    kn, ka, kb = tk.StringVar(), tk.StringVar(), tk.StringVar(value=datetime.date.today().isoformat())
    for i, (t, v, w) in enumerate((("Nama kegiatan", kn, 38), ("Anggaran (Rp)", ka, 18), ("Tanggal buat (YYYY-MM-DD)", kb, 18))):
        ttk.Label(g, text=t, style="Card.TLabel").grid(row=0, column=i, sticky="w", padx=10, pady=(8, 0))
        ttk.Entry(g, textvariable=v, width=w).grid(row=1, column=i, padx=10, pady=(0, 10))

    def kadd():
        try:
            a = int("".join(c for c in ka.get() if c.isdigit()))
            d = datetime.date.fromisoformat(kb.get().strip())
            assert a > 0 and kn.get().strip() and kn.get().strip() not in [k["nama"] for k in keg]
        except Exception:
            messagebox.showwarning("Periksa isian", "Nama harus unik, anggaran angka, tanggal format YYYY-MM-DD.")
            return
        keg.append({"id": int(datetime.datetime.now().timestamp() * 1000), "nama": kn.get().strip(), "anggaran": a, "buat": d.isoformat()})
        save(data, keg)
        kn.set("")
        ka.set("")
        refresh()
    btn(g, "+  Tambah Kegiatan", kadd, TEAL).grid(row=1, column=3, padx=10, pady=(0, 10))

    kcols = ("nama", "buat", "ang", "pakai", "sisa", "pct")
    kh = (("Kegiatan", 260), ("Tgl Buat", 90), ("Anggaran", 120), ("Terpakai", 120), ("Sisa", 120), ("%", 60))
    dcols = ("nama", "buat", "selesai", "ang", "pakai", "sisa", "pct")
    dh = (("Kegiatan", 240), ("Tgl Buat", 90), ("Tgl Selesai", 90), ("Anggaran", 115), ("Terpakai", 115), ("Sisa", 115), ("%", 55))
    h_wait, kt_wait = mk(tab2, "", "Wait.Treeview", kcols, kh, {"over": RED}, 6)
    h_done, kt_done = mk(tab2, "", "Done.Treeview", dcols, dh, {"over": RED}, 6)

    def pilih(t):
        return [k for k in keg if str(k["id"]) in t.selection()]

    def tandai():
        for k in pilih(kt_wait):
            k["selesai"] = datetime.date.today().isoformat()
        save(data, keg)
        refresh()

    def buka():
        for k in pilih(kt_done):
            if k.get("selesai"):
                del k["selesai"]
            else:
                messagebox.showinfo("Info", f"'{k['nama']}' selesai karena anggaran sudah terpakai 100%.\nHapus transaksinya atau tambah anggaran untuk membukanya kembali.")
        save(data, keg)
        refresh()

    def khapus(t):
        sel = pilih(t)
        if sel and messagebox.askyesno("Hapus", "Hapus kegiatan terpilih? Transaksi tetap tersimpan."):
            keg[:] = [k for k in keg if k not in sel]
            save(data, keg)
            refresh()
    kb2 = tk.Frame(tab2, bg=BG)
    kb2.pack(fill="x")
    btn(kb2, "✔ Tandai Selesai (tabel atas)", tandai, GREEN).pack(side="left", padx=(0, 8))
    btn(kb2, "↩ Buka Kembali (tabel bawah)", buka, AMBER).pack(side="left", padx=(0, 8))
    btn(kb2, "Hapus (atas)", lambda: khapus(kt_wait), RED).pack(side="left", padx=(0, 8))
    btn(kb2, "Hapus (bawah)", lambda: khapus(kt_done), SLATE).pack(side="left")

    def refresh():
        def opts(cur, vals):
            vals = ["Semua"] + sorted(set(vals))
            if cur.get() not in vals:
                cur.set("Semua")
            return vals
        cbs["Bulan"]["values"] = opts(bln, [t["tgl"][:7] for t in data])
        cbs["Kategori"]["values"] = opts(fk, [t["kat"] for t in data])
        cbs["Kegiatan"]["values"] = opts(fg, [k["nama"] for k in keg] + [t["keg"] for t in data if t.get("keg")])
        v = view()
        for tv, jn in ((tv_in, "Pemasukan"), (tv_out, "Pengeluaran")):
            tv.delete(*tv.get_children())
            rows = sorted((x for x in v if x["jenis"] == jn), key=lambda t: (t["tgl"], t["id"]), reverse=True)
            for i, t in enumerate(rows):
                tv.insert("", "end", iid=str(t["id"]), tags=("r", "odd") if i % 2 else ("r",),
                          values=(t["tgl"], t.get("nb", ""), t["kat"], t.get("keg", ""), t["ket"], rp(t["jml"])))
        ti = sum(t["jml"] for t in v if t["jenis"] == "Pemasukan")
        to = sum(t["jml"] for t in v if t["jenis"] == "Pengeluaran")
        val["in"].config(text=rp(ti))
        val["out"].config(text=rp(to))
        val["saldo"].config(text=rp(ti - to))
        lf_in.config(text=f"⬇  Tabel Pemasukan   •   total {rp(ti)}", fg=GREEN)
        lf_out.config(text=f"⬆  Tabel Pengeluaran   •   total {rp(to)}", fg=RED)

        stat = [(k, *keg_status(k, data)) for k in keg]
        avail = [k["nama"] for k, ok, _ in stat if not ok]  # kegiatan selesai/100% tidak muncul saat input
        cb_kg["values"] = avail
        if kg.get() not in avail:
            kg.set("")
        kt_wait.delete(*kt_wait.get_children())
        kt_done.delete(*kt_done.get_children())
        nw = nd = 0
        for k, ok, tsel in stat:
            r = real(data, k["nama"])
            pct = round(r / k["anggaran"] * 100) if k["anggaran"] else 0
            base = (k["nama"], tgl_buat(k)) + ((tsel,) if ok else ()) + (rp(k["anggaran"]), rp(r), rp(k["anggaran"] - r), f"{pct}%")
            tg = ("over",) if r > k["anggaran"] else ()
            if ok:
                kt_done.insert("", "end", iid=str(k["id"]), values=base, tags=tg)
                nd += 1
            else:
                kt_wait.insert("", "end", iid=str(k["id"]), values=base, tags=tg)
                nw += 1
        h_wait.config(text=f"⏳  Kegiatan Belum Selesai  ({nw})", fg="#b45309")
        h_done.config(text=f"✅  Kegiatan Selesai  ({nd})", fg=BLUE)

    for c in cbs.values():
        c.bind("<<ComboboxSelected>>", lambda e: refresh())
    set_kat()
    refresh()
    root.mainloop()


if __name__ == "__main__":
    main()
