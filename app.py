"""
Kalite Departmanı – Streamlit Uygulaması (v4) – Makine/Ürün/Kalite + İş Emri Entegrasyonları + Isı Haritası + Trendler
================================================
Bu uygulama, kalite departmanının Excel tablosunu Streamlit arayüzünde analiz eder,
opsiyonel kurallar ile doğrulama yapar ve rapor üretir.

KULLANIM:
- Terminal:  streamlit run app.py
- Gerekli paketler için aşağıdaki requirements içeriğini kullanın.

NOT:
- Excel dosyanızdaki sayfa adları serbesttir.
- Opsiyonel olarak bir "Kurallar" sayfası oluşturursanız, aşağıdaki şemayı kullanın:

  | Kolon | KuralTürü | Parametre |
  |-------|-----------|-----------|
  | SKU   | regex     | ^[A-Z0-9_-]{4,20}$ |
  | LotNo | zorunlu   |           |
  | Miktar| minmax    | 0;100000  |
  | Durum | set       | UYGUN;RED;ŞARTLI |

KuralTürü değerleri: "regex", "zorunlu", "minmax" ("min;max"), "set" (noktalı virgülle liste)

"""

from __future__ import annotations
import io
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ====== YARDIMCI ======
@dataclass
class Rule:
    column: str
    rtype: str  # regex | zorunlu | minmax | set
    param: str


def parse_rules(df_rules: pd.DataFrame) -> List[Rule]:
    rules: List[Rule] = []
    if df_rules is None or df_rules.empty:
        return rules
    # Kolon adlarını normalize et
    cols_map = {c.lower(): c for c in df_rules.columns}
    col_col = cols_map.get("kolon") or cols_map.get("column")
    type_col = cols_map.get("kuraltürü") or cols_map.get("kural_türü") or cols_map.get("ruletype")
    param_col = cols_map.get("parametre") or cols_map.get("param")
    if not (col_col and type_col and param_col):
        return rules
    for _, r in df_rules.iterrows():
        col = str(r[col_col]).strip()
        rtype = str(r[type_col]).strip().lower()
        param = "" if pd.isna(r[param_col]) else str(r[param_col]).strip()
        if col:
            rules.append(Rule(col, rtype, param))
    return rules


def apply_rules(df: pd.DataFrame, rules: List[Rule]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Kuralları uygular, iki DataFrame döner: (hatalar, özet)."""
    if df is None or df.empty or not rules:
        return pd.DataFrame(), pd.DataFrame()

    errors = []
    for rule in rules:
        if rule.column not in df.columns:
            errors.append({
                "Sayfa": st.session_state.get("active_sheet", "<seçilmedi>"),
                "Kolon": rule.column,
                "Kural": rule.rtype,
                "Hata": "Kolon bulunamadı"
            })
            continue
        series = df[rule.column]
        if rule.rtype == "zorunlu":
            mask = series.isna() | (series.astype(str).str.strip() == "")
            idx = df.index[mask]
            for i in idx:
                errors.append({"Sayfa": st.session_state.get("active_sheet", ""), "Kolon": rule.column, "Kural": "zorunlu", "Hata": "Boş bırakılamaz", "Satır": int(i)})
        elif rule.rtype == "regex":
            try:
                mask = ~series.astype(str).str.match(rule.param, na=False)
            except Exception:
                mask = pd.Series(False, index=series.index)
            idx = df.index[mask]
            for i in idx:
                errors.append({"Sayfa": st.session_state.get("active_sheet", ""), "Kolon": rule.column, "Kural": f"regex: {rule.param}", "Hata": "Desen uyumsuz", "Satır": int(i)})
        elif rule.rtype == "minmax":
            try:
                parts = [p.strip() for p in rule.param.split(";")]
                min_v = float(parts[0]) if parts[0] != "" else -np.inf
                max_v = float(parts[1]) if len(parts) > 1 and parts[1] != "" else np.inf
            except Exception:
                min_v, max_v = -np.inf, np.inf
            with pd.option_context('mode.use_inf_as_na', True):
                s = pd.to_numeric(series, errors='coerce')
            mask = (s < min_v) | (s > max_v)
            idx = df.index[mask.fillna(True)]
            for i in idx:
                errors.append({"Sayfa": st.session_state.get("active_sheet", ""), "Kolon": rule.column, "Kural": f"minmax: {min_v};{max_v}", "Hata": "Aralık dışı", "Satır": int(i)})
        elif rule.rtype == "set":
            allowed = {p.strip() for p in rule.param.split(";") if p.strip() != ""}
            mask = ~series.astype(str).isin(allowed)
            idx = df.index[mask]
            for i in idx:
                errors.append({"Sayfa": st.session_state.get("active_sheet", ""), "Kolon": rule.column, "Kural": f"set: {sorted(allowed)}", "Hata": "İzinli değer değil", "Satır": int(i)})
        else:
            errors.append({"Sayfa": st.session_state.get("active_sheet", ""), "Kolon": rule.column, "Kural": rule.rtype, "Hata": "Bilinmeyen kural türü"})

    df_err = pd.DataFrame(errors)
    summary = (df_err.groupby(["Sayfa", "Kolon", "Kural", "Hata"]).size().reset_index(name="Adet")
               if not df_err.empty else pd.DataFrame(columns=["Sayfa", "Kolon", "Kural", "Hata", "Adet"]))
    return df_err, summary


def profile_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    prof = []
    for c in df.columns:
        s = df[c]
        prof.append({
            "Kolon": c,
            "Tip": str(s.dtype),
            "Satır": len(s),
            "Boş": int(s.isna().sum()),
            "Benzersiz": int(s.nunique(dropna=True)),
        })
    return pd.DataFrame(prof)


# ====== UI ======
st.set_page_config(page_title="Kalite – Excel Doğrulama", layout="wide")
st.title("📊 Kalite Departmanı – Excel Doğrulama Uygulaması")
st.caption("Excel yükleyin, sayfaları inceleyin, kuralları uygulayın, makine/ürün bazında kalite KPI'larını görün ve rapor indirin.")

with st.sidebar:
    st.header("Ayarlar")
    st.session_state.setdefault("tz", "Europe/Istanbul")
    st.write("Zaman Dilimi:", st.session_state["tz"])
    st.markdown("**1) Ana Excel (QC MASTER) Yükle**")
    file = st.file_uploader("QC Master Excel (.xlsx)", type=["xlsx"], key="qc_master")
    st.markdown("**2) Canias Gramaj Form(lar)ı (opsiyonel)**")
grams_files = st.file_uploader("Bir veya daha fazla Excel/PDF'den çevrilmiş tablo (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_grams")
tol = st.number_input("Gramaj toleransı (±)", min_value=0.0, max_value=50.0, value=5.0, step=0.5)

st.markdown("**3) Diğer Alt Formlar (opsiyonel)**")
uygun_terms = st.text_input("'Uygun' sayılacak ifadeler (virgülle)", value="Uygun,OK,Pass")
red_terms = st.text_input("'Uygun Değil'/'Hata' sayılacak ifadeler (virgülle)", value="Uygun Değil,NOK,Fail,Red,Hata")

sensory_files = st.file_uploader("Duyusal formlar (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_duyusal")
pack_files    = st.file_uploader("Ambalaj/Etiket formları (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_ambalaj")
cap_files     = st.file_uploader("Kapak formları (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_kapak")
inner_files   = st.file_uploader("Koli/Kutu İçi Adet formları (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_koli_ici")
gas_files     = st.file_uploader("Gaz kontrol formları (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="canias_gaz")

auto_profile = st.toggle("Yükler yüklemez profil çıkar", value=True)
show_raw = st.toggle("Ham veriyi göster", value=True)", type=["xlsx"])
    auto_profile = st.toggle("Yükler yüklemez profil çıkar", value=True)
    show_raw = st.toggle("Ham veriyi göster", value=True)

# ==== ŞEMA HARİTALAMA (Makine/Ürün/Kalite) ====
def map_columns_ui(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    st.markdown("### 🧭 Kolon Haritalama")
    st.caption("Makine, Ürün, Sonuç veya parametre alanlarının Excel'deki karşılıklarını seçin.")
    cols = [None] + list(df.columns)

    def pick(label, candidates):
        guess = None
        lower = {c.lower(): c for c in df.columns}
        for cand in candidates:
            if cand in lower:
                guess = lower[cand]
                break
        return st.selectbox(label, options=cols, index=(cols.index(guess) if guess in cols else 0), key=f"map_{label}")

    mapping = {
        "job": pick("İş Emri ID kolonu", ["iş emri id", "iş emri no", "is emri id", "is emri no", "iş emri", "is emri", "ie no", "ie_no", "ie numara", "ie numarası"]),
        "date": pick("Tarih kolonu", ["tarih", "kontrol tarih", "kontrol tarihi", "date"]),
        "machine": pick("Makine kolonu", ["makine", "makine adı", "makine adi", "machine"]),
        "product": pick("Ürün kolonu", ["ürün", "urun", "ürün adı", "urun adi", "product", "malzeme adı", "malzeme adi"]),
        "qty": pick("Üretim miktarı (adet) kolonu", ["üretim miktarı adet", "uretim miktari adet", "miktar", "adet", "quantity", "qty"]),
        "check_count": pick("Kontrol sayısı kolonu (ops.)", ["kontrol sayısı", "kontrol sayisi", "num kontrol", "kontrol adet"]),
        # Parametre alanları
        "p_duyusal": pick("DUYUSAL kolonu", ["duyusal"]),
        "p_gramaj": pick("GRAMAJ kolonu", ["gramaj"]),
        "p_ambalaj": pick("AMBALAJ/KOLİ ETİKET kolonu", ["ambalaj/koli etiket", "ambalaj", "etiket"]),
        "p_kapak": pick("KAPAK kolonu", ["kapak"]),
        "p_koli_ici": pick("KOLİ/KUTU İÇİ ADET kolonu", ["koli/kutu içi adet", "kutu içi adet", "koli içi adet"]),
        "p_gaz": pick("GAZ KONTROL kolonu", ["gaz kontrol", "gaz"]),
        "p_total": pick("TOPLAM HATA kolonu", ["toplam hata", "toplam_hata"]),
    }
    return mapping

if file:
    try:
        xl = pd.ExcelFile(file, engine="openpyxl")
        sheet_names = xl.sheet_names
    except Exception as e:
        st.error(f"QC Master okunamadı: {e}")
        st.stop()

    st.success(f"QC Master içinde {len(sheet_names)} sayfa bulundu: {', '.join(sheet_names)}")

    colL, colR = st.columns([2, 1])
    with colL:
        active_sheet = st.selectbox("Sayfa seçin", options=sheet_names, index=0)
        st.session_state["active_sheet"] = active_sheet
    with colR:
        rules_sheet_name = st.text_input("Kurallar sayfa adı (opsiyonel)", value="Kurallar")
        run_checks = st.button("✅ Kuralları Uygula")

    try:
        df_active = pd.read_excel(file, sheet_name=active_sheet, engine="openpyxl")
    except Exception as e:
        st.error(f"Seçili sayfa okunamadı: {e}")
        st.stop()
    except Exception as e:
        st.error(f"Seçili sayfa okunamadı: {e}")
        st.stop()

    # Haritalama UI
    mapping = map_columns_ui(df_active)

    # Filtreler
    with st.expander("🔎 Filtreler", expanded=False):
        machines = sorted(df_active[mapping["machine"]].dropna().unique()) if mapping["machine"] else []
        products = sorted(df_active[mapping["product"]].dropna().unique()) if mapping["product"] else []
        m_sel = st.multiselect("Makine", machines)
        p_sel = st.multiselect("Ürün", products)
        date_range = None
        if mapping["date"]:
            min_d = pd.to_datetime(df_active[mapping["date"]], errors='coerce').min()
            max_d = pd.to_datetime(df_active[mapping["date"]], errors='coerce').max()
            if pd.notna(min_d) and pd.notna(max_d):
                date_range = st.date_input("Tarih aralığı", value=(min_d.date(), max_d.date()))

    # Filtre uygula
    df_view = df_active.copy()
    if mapping["machine"] and m_sel:
        df_view = df_view[df_view[mapping["machine"].strip()].isin(m_sel)]
    if mapping["product"] and p_sel:
        df_view = df_view[df_view[mapping["product"].strip()].isin(p_sel)]
    if mapping["date"] and date_range and len(date_range) == 2:
        d0, d1 = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        s = pd.to_datetime(df_view[mapping["date"]], errors='coerce')
        df_view = df_view[(s >= d0) & (s <= d1)]

    tabs = st.tabs(["Veri", "Profil", "Kalite KPI", "Kurallar & Hata Listesi", "Özet Rapor", "İndir"])

    with tabs[0]:
        st.subheader(f"📄 Veri – {active_sheet}")
        st.dataframe(df_view, use_container_width=True)

    with tabs[1]:
        st.subheader("🧪 Profil")
        if auto_profile:
            prof = profile_df(df_view)
            st.dataframe(prof, use_container_width=True)
        else:
            if st.button("Profil Çıkar"):
                prof = profile_df(df_view)
                st.dataframe(prof, use_container_width=True)

    # ===== Canias GRAMAJ entegrasyonu =====
st.markdown("### 🔗 Canias Entegrasyonları (İş Emri bazlı)")
mapping = map_columns_ui(df_active)

# Yardımcı: çoklu dosyayı oku
def read_many(files):
    out = []
    for f in files or []:
        try:
            out.append(pd.read_excel(f, engine="openpyxl"))
        except Exception as e:
            st.warning(f"Dosya okunamadı: {getattr(f,'name','dosya')}: {e}")
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame()

# Yardımcı: kolon adı tahmini
ndefaults = lambda cols: {c.lower(): c for c in cols}

def norm_map_grams(df):
    if df.empty: return {"job":None,"gram":None,"target":None,"tarih":None}
    L = ndefaults(df.columns)
    return {
        "job": L.get("iş emri no") or L.get("is emri no") or L.get("iş emri id") or L.get("is emri id") or L.get("ie no") or L.get("iş emri") or L.get("is emri"),
        "gram": L.get("gram") or L.get("net gram") or L.get("olcum") or L.get("ölçüm"),
        "target": L.get("malzeme kartı gram") or L.get("malzeme karti gram") or L.get("hedef gram") or L.get("target"),
        "tarih": L.get("kontrol tarih") or L.get("kontrol tarihi") or L.get("tarih"),
    }

def norm_map_yesno(df):
    if df.empty: return {"job":None, "flag":None, "status":None}
    L = ndefaults(df.columns)
    # Esnek: açık bir Hata/Flag sütunu varsa onu kullan; yoksa Uygunluk metni
    return {
        "job": L.get("iş emri no") or L.get("is emri no") or L.get("iş emri id") or L.get("is emri id") or L.get("ie no") or L.get("iş emri") or L.get("is emri"),
        "flag": L.get("hata") or L.get("hata sayısı") or L.get("hata sayisi") or L.get("uygunsuz") or L.get("fail") or L.get("nok"),
        "status": L.get("uygunluk") or L.get("sonuç") or L.get("sonuc") or L.get("durum")
    }

# 1) GRAMAJ
grams_df = read_many(grams_files)
if not grams_df.empty and mapping["job"]:
    gm = norm_map_grams(grams_df)
    if gm["job"] and gm["gram"] and gm["target"]:
        gd = grams_df[[gm["job"], gm["gram"], gm["target"]]].copy()
        gd.columns = ["JOB", "GRAM", "TARGET"]
        gd["deviation"] = (pd.to_numeric(gd["GRAM"], errors='coerce') - pd.to_numeric(gd["TARGET"], errors='coerce')).abs()
        gd["out"] = gd["deviation"] > tol
        gram_summary = gd.groupby("JOB").agg(kontrol_sayisi=("GRAM", "count"), gramaj_hata=("out", "any")).reset_index()
        gram_summary["gramaj_hata"] = gram_summary["gramaj_hata"].astype(int)
        df_active = df_active.copy()
        df_active["JOB"] = df_active[mapping["job"]].astype(str)
        merged = df_active.merge(gram_summary, on="JOB", how="left")
        # GRAMAJ kolonunu 0/1 olarak güncelle
        if mapping["p_gramaj"]:
            pg = mapping["p_gramaj"]
            merged[pg] = merged[pg].fillna(0)
            merged.loc[merged["gramaj_hata"].fillna(0).astype(int) == 1, pg] = 1
        # Kontrol sayısı boşsa Canias'tan al
        if mapping["check_count"]:
            cc = mapping["check_count"]
            merged[cc] = merged[cc].where(merged[cc].fillna(0) > 0, merged["kontrol_sayisi"]) 
        df_active = merged.drop(columns=[c for c in ["JOB", "kontrol_sayisi", "gramaj_hata"] if c in merged.columns])

# 2) DUYUSAL / AMBALAJ / KAPAK / KOLİİÇİ / GAZ  → esnek kurallar
uyguns = {s.strip().lower() for s in (uygun_terms or "").split(',') if s.strip()}
reds   = {s.strip().lower() for s in (red_terms or "").split(',') if s.strip()}

def merge_yesno(active, files, target_col):
    if not files or not mapping["job"] or not mapping[target_col]:
        return active
    df = read_many(files)
    if df.empty:
        return active
    nm = norm_map_yesno(df)
    if not nm["job"]:
        return active
    temp = df[[nm["job"]]].copy()
    temp.columns = ["JOB"]
    # Flag öncelikli
    if nm["flag"] and nm["flag"] in df.columns:
        flag = pd.to_numeric(df[nm["flag"]], errors='coerce').fillna(0)
        df2 = pd.DataFrame({"JOB": df[nm["job"]].astype(str), "err": (flag > 0).astype(int)})
    elif nm["status"] and nm["status"] in df.columns:
        stxt = df[nm["status"]].astype(str).str.strip().str.lower()
        df2 = pd.DataFrame({"JOB": df[nm["job"]].astype(str), "err": stxt.apply(lambda x: 1 if (x in reds) else (0 if (x in uyguns) else 0))})
    else:
        # Kolonlar yoksa, bu formda kayıt olması "inceleme yapıldı" sayılır ama hata bilinmez → 0
        df2 = pd.DataFrame({"JOB": df[nm["job"]].astype(str), "err": 0})
    agg = df2.groupby("JOB").agg(err=("err","max")).reset_index()
    active = active.copy()
    active["JOB"] = active[mapping["job"]].astype(str)
    merged = active.merge(agg, on="JOB", how="left")
    col = mapping[target_col]
    merged[col] = merged[col].fillna(0)
    merged.loc[merged["err"].fillna(0).astype(int) == 1, col] = 1
    return merged.drop(columns=[c for c in ["JOB","err"] if c in merged.columns])

for files, key in [
    (sensory_files, "p_duyusal"),
    (pack_files,    "p_ambalaj"),
    (cap_files,     "p_kapak"),
    (inner_files,   "p_koli_ici"),
    (gas_files,     "p_gaz"),
]:
    df_active = merge_yesno(df_active, files, key)

    # ===== KPI Sekmeleri =====
    tabs = st.tabs(["Veri", "Profil", "Kalite KPI", "Isı Haritası", "Trendler", "Kurallar & Hata Listesi", "Özet Rapor", "İndir"])

    with tabs[0]:
        st.subheader(f"📄 Veri – {active_sheet}")
        st.dataframe(df_active, use_container_width=True)

    with tabs[1]:
        st.subheader("🧪 Profil")
        if auto_profile:
            prof = profile_df(df_active)
            st.dataframe(prof, use_container_width=True)
        else:
            if st.button("Profil Çıkar"):
                prof = profile_df(df_active)
                st.dataframe(prof, use_container_width=True)
        else:
            if st.button("Profil Çıkar"):
                prof = profile_df(df_active)
                st.dataframe(prof, use_container_width=True)

    # Sonuç normalizasyonu ve hesaplamalar
    def infer_result_from_params(row, mapping):
        params = []
        for k in ["p_duyusal", "p_gramaj", "p_ambalaj", "p_kapak", "p_koli_ici", "p_gaz"]:
            col = mapping.get(k)
            if col and col in row.index:
                try:
                    params.append(float(row[col]))
                except Exception:
                    params.append(0.0)
        total_err = sum(1 for v in params if (pd.notna(v) and float(v) > 0))  # parametre bazlı 0/1 kabul
        # Eğer TOPLAM HATA kolonu varsa onu kullan, yoksa parametreden üret
        total_col = mapping.get("p_total")
        if total_col and total_col in row.index and pd.notna(row[total_col]):
            th = float(row[total_col])
        else:
            th = total_err
        if th <= 0:
            return "UYGUN"
        elif th <= 3:
            return "ŞARTLI"
        return "RED"

    # KPI hesapla
    df_kpi = df_active.copy()
    if mapping["qty"] and mapping["machine"] and mapping["product"]:
        df_kpi["__sonuc"] = df_kpi.apply(lambda r: infer_result_from_params(r, mapping), axis=1)
        qty = pd.to_numeric(df_kpi[mapping["qty"]], errors='coerce').fillna(1)
        total = int(qty.sum()) if len(qty) else 0
        ok = int(qty[df_kpi["__sonuc"] == "UYGUN"].sum()) if total else 0
        red = int(qty[df_kpi["__sonuc"] == "RED"].sum()) if total else 0
        cond = int(qty[df_kpi["__sonuc"] == "ŞARTLI"].sum()) if total else 0
        scrap_rate = (red / total * 100) if total else 0.0
        pass_rate = (ok / total * 100) if total else 0.0
        ppm = (red / total * 1_000_000) if total else 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Toplam Üretim (adet)", f"{total:,}")
        c2.metric("Geçiş Oranı (Pass %)", f"{pass_rate:.2f}%")
        c3.metric("Hurda Oranı (Scrap %)", f"{scrap_rate:.2f}%")
        c4.metric("RED PPM", f"{ppm:,.0f}")

        st.markdown("#### 🏭 Makine Bazlı Sonuç")
        g_m = df_kpi.groupby(mapping["machine"], dropna=True).apply(
            lambda d: pd.Series({
                "Toplam": int(pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum()),
                "Pass%": ((pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1)[d["__sonuc"]=="UYGUN"].sum() / (pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum())) * 100),
                "Scrap%": ((pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1)[d["__sonuc"]=="RED"].sum() / (pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum())) * 100),
            })
        ).reset_index().rename(columns={mapping["machine"]: "Makine"})
        st.dataframe(g_m.sort_values("Scrap%", ascending=False), use_container_width=True)

        st.markdown("#### 📦 Ürün Bazlı Sonuç")
        g_p = df_kpi.groupby(mapping["product"], dropna=True).apply(
            lambda d: pd.Series({
                "Toplam": int(pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum()),
                "Pass%": ((pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1)[d["__sonuc"]=="UYGUN"].sum() / (pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum())) * 100),
                "Scrap%": ((pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1)[d["__sonuc"]=="RED"].sum() / (pd.to_numeric(d[mapping["qty"]], errors='coerce').fillna(1).sum())) * 100),
            })
        ).reset_index().rename(columns={mapping["product"]: "Ürün"})
        st.dataframe(g_p.sort_values("Scrap%", ascending=False), use_container_width=True)


    with tabs[3]:
        st.subheader("🗺️ Isı Haritası – Makine × Ürün (RED adet)")
        if mapping["machine"] and mapping["product"]:
            temp = df_active.copy()
            temp["__sonuc"] = temp.apply(lambda r: infer_result_from_params(r, mapping), axis=1)
            temp["__red_adet"] = 1
            if mapping["qty"]:
                temp["__red_adet"] = pd.to_numeric(temp[mapping["qty"]], errors='coerce').fillna(1)
            temp = temp[temp["__sonuc"]=="RED"]
            pivot = temp.pivot_table(index=mapping["machine"], columns=mapping["product"], values="__red_adet", aggfunc="sum", fill_value=0)
            st.dataframe(pivot, use_container_width=True)
        else:
            st.info("Isı haritası için Makine ve Ürün kolonlarını haritalayın.")

    with tabs[4]:
        st.subheader("📉 Trendler – Günlük RED ve RED PPM")
        if mapping["date"]:
            tdf = df_active.copy()
            tdf["__sonuc"] = tdf.apply(lambda r: infer_result_from_params(r, mapping), axis=1)
            tdf["__qty"] = pd.to_numeric(tdf[mapping["qty"]], errors='coerce').fillna(1) if mapping["qty"] else 1
            tdf["__date"] = pd.to_datetime(tdf[mapping["date"]], errors='coerce').dt.date
            grp = tdf.groupby("__date").agg(total=("__qty","sum"), red=("__qty", lambda s: s[tdf.loc[s.index,"__sonuc"]=="RED"].sum())).reset_index()
            grp["ppm"] = grp.apply(lambda r: (r["red"]/r["total"]*1_000_000) if r["total"] else 0, axis=1)
            import altair as alt
            if not grp.empty:
                line1 = alt.Chart(grp).mark_line().encode(x="__date:T", y="red:Q").properties(height=250, title="Günlük RED (adet)")
                line2 = alt.Chart(grp).mark_line().encode(x="__date:T", y="ppm:Q").properties(height=250, title="Günlük RED PPM")
                st.altair_chart(line1, use_container_width=True)
                st.altair_chart(line2, use_container_width=True)
            else:
                st.info("Trend grafiği için geçerli tarih ve miktar verisi bulunamadı.")
        else:
            st.info("Trend grafikleri için Tarih kolonu haritalayın.")

    with tabs[3]:
        st.subheader("🧩 Kurallar & Hatalar")
        df_rules = pd.DataFrame()
        if rules_sheet_name in sheet_names:
            try:
                df_rules = pd.read_excel(file, sheet_name=rules_sheet_name, engine="openpyxl")
                st.markdown("**Yüklenen Kurallar**")
                st.dataframe(df_rules, use_container_width=True, height=200)
            except Exception as e:
                st.warning(f"Kurallar sayfası okunamadı: {e}")
        else:
            st.info("Kurallar sayfası bulunamadı. İsterseniz bir 'Kurallar' sayfası ekleyin.")

        rules = parse_rules(df_rules) if not df_rules.empty else []
        if run_checks and rules:
            errs, summary = apply_rules(df_active, rules)
            st.session_state["err_df"] = errs
            st.session_state["sum_df"] = summary
        if "err_df" in st.session_state and not st.session_state["err_df"].empty:
            st.error(f"Toplam hata: {len(st.session_state['err_df'])}")
            st.dataframe(st.session_state["err_df"], use_container_width=True)
        else:
            st.success("Hata bulunamadı ya da kurallar uygulanmadı.")

    with tabs[4]:
        st.subheader("📈 Özet Rapor")
        if "sum_df" in st.session_state and not st.session_state["sum_df"].empty:
            st.dataframe(st.session_state["sum_df"], use_container_width=True)
        else:
            st.info("Özet rapor için önce kuralları uygulayın.")

    with tabs[5]:
        st.subheader("💾 İndir")
        # Görünüm CSV
        csv = df_view.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Filtrelenmiş veriyi CSV indir", data=csv, file_name=f"{active_sheet}_filtreli.csv", mime="text/csv")

        # Hata ve özet excel
        if "err_df" in st.session_state or "sum_df" in st.session_state:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                if "err_df" in st.session_state and not st.session_state["err_df"].empty:
                    st.session_state["err_df"].to_excel(writer, index=False, sheet_name="Hatalar")
                if "sum_df" in st.session_state and not st.session_state["sum_df"].empty:
                    st.session_state["sum_df"].to_excel(writer, index=False, sheet_name="Ozet")
            st.download_button(
                "Hata & Özet (XLSX) indir",
                data=buffer.getvalue(),
                file_name="kalite_raporu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Başlamak için sol menüden bir Excel yükleyin.")

# ====== requirements.txt (referans) ======
# streamlit==1.37.0
# pandas>=2.0.0
# openpyxl>=3.1.2
# numpy>=1.25.0
