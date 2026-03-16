import re
import math
import pandas as pd
import streamlit as st
import plotly.express as px
import streamlit.components.v1 as components
import base64
from pathlib import Path
from urllib.parse import quote

# -----------------------------
# Helpers
# -----------------------------
BILDE_RE = re.compile(r"lodd\s*bilde\s*([A-Z])(?:\s|-|$)", re.IGNORECASE)

def find_old_header_row(raw: pd.DataFrame) -> int | None:
    """Find header row for old Vipps export (contains 'Salgssted')."""
    for i in range(min(len(raw), 200)):
        row = raw.iloc[i].astype(str)
        if row.str.contains(r"\bSalgssted\b", case=False, na=False).any():
            return i
    return None


def read_vipps_report(uploaded_file) -> pd.DataFrame:
    """
    Read Vipps report in either old or new format and return a normalized dataframe.

    Internal normalized columns:
    - LotteriKategori   <- old: Salgssted / new: Kategori
    - Transaksjonstype  <- old: Transaksjonstype / new: Type
    - Brutto            <- old: Brutto / new: Beløp
    - Navn              <- new: Kundens navn (or existing Navn)
    """
    xls = pd.ExcelFile(uploaded_file)

    preferred_sheets = sorted(
        xls.sheet_names,
        key=lambda s: (
            0 if str(s).strip().lower() == "detaljer" else
            1 if "detalj" in str(s).lower() else
            2 if "rapport" in str(s).lower() else
            3
        )
    )

    for sheet_name in preferred_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.columns = [str(c).strip() for c in df.columns]

            # NEW FORMAT
            required_new = {"Beløp", "Kundens navn", "Kategori"}
            if required_new.issubset(df.columns):
                df = df.copy()
                df["Brutto"] = df["Beløp"]
                df["Navn"] = df["Kundens navn"]
                df["LotteriKategori"] = df["Kategori"]

                if "Type" in df.columns:
                    df["Transaksjonstype"] = df["Type"]
                elif "Transaksjonstype" not in df.columns:
                    df["Transaksjonstype"] = ""

                return df

            # OLD FORMAT
            required_old = {"Salgssted", "Brutto"}
            if required_old.issubset(df.columns):
                df = df.copy()
                df["LotteriKategori"] = df["Salgssted"]

                if "Transaksjonstype" not in df.columns:
                    df["Transaksjonstype"] = ""

                return df

        except Exception:
            pass

        try:
            raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            header_row = find_old_header_row(raw)

            if header_row is not None:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

                if len(df) > 0 and str(df.iloc[0].get("Salgssted", "")).strip().lower() == "salgssted":
                    df.columns = df.iloc[0].tolist()
                    df = df.iloc[1:].copy()

                df.columns = [str(c).strip() for c in df.columns]
                df = df.copy()
                df["LotteriKategori"] = df["Salgssted"]

                if "Transaksjonstype" not in df.columns:
                    df["Transaksjonstype"] = ""

                return df
        except Exception:
            pass

    raise ValueError(
        "Fant ikke et støttet rapportformat. "
        "Forventet enten gammel Vipps-rapport eller nytt format med "
        "'Beløp', 'Kundens navn' og 'Kategori'."
    )

def build_full_name(row) -> str:
    # New format
    navn = str(row.get("Navn", "")).strip()
    if navn and navn.lower() != "nan":
        return navn

    # Old format
    fn = str(row.get("Fornavn", "")).strip()
    en = str(row.get("Etternavn", "")).strip()

    if fn and fn.lower() != "nan":
        return (fn + " " + en).strip() if en and en.lower() != "nan" else fn

    # Fallback
    msg = str(row.get("Melding", "")).strip()
    if msg and msg.lower() != "nan":
        return msg

    return "Ukjent"

def extract_bilde(salgssted: str) -> str | None:
    if not salgssted or str(salgssted).lower() == "nan":
        return None
    m = BILDE_RE.search(str(salgssted))
    return m.group(1).upper() if m else None

def as_int_floor(x) -> int:
    try:
        return int(math.floor(float(x)))
    except Exception:
        return 0


# -----------------------------
# UI
# -----------------------------
st.set_page_config(
    page_title="Kunstlotteri – NHO Kunst og Kultur 🎨",
    page_icon="🎨",
    layout="wide",
)

st.title("Kunstlotteri – NHO Kunst og Kultur")
# st.caption("Last opp Vipps-rapporten. Velg bilde. Kopiér deltakerlisten og lim inn i Wheel of Names.")

uploaded = st.file_uploader("Last opp Vipps-rapporten (.xlsx)", type=["xlsx"])

with st.expander("Innstillinger", expanded=False):
    loddpris = st.number_input("Loddpris (kr)", min_value=1, value=20, step=1)
    name_mode = st.radio("Navnformat", ["Fullt navn", "Kun fornavn"], horizontal=True)
    round_down = st.checkbox("Rund ned til heltall lodd (anbefalt)", value=True)

if not uploaded:
    st.info("Last opp Vipps-rapporten for å komme i gang.")
    st.stop()

try:
    df = read_vipps_report(uploaded)
except Exception as e:
    st.error(str(e))
    st.stop()

required_cols = {"LotteriKategori", "Transaksjonstype", "Brutto"}
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Mangler forventede kolonner i rapporten: {missing}")
    st.stop()

# Normalize & filter relevant rows
df["LotteriKategori"] = df["LotteriKategori"].astype(str).str.strip()
df["Transaksjonstype"] = df["Transaksjonstype"].astype(str).str.strip().str.lower()

df_lodd = df[
    df["LotteriKategori"].str.contains(r"lodd\s*bilde", case=False, na=False)
    & df["Transaksjonstype"].isin(["salg", "belastning", ""])
].copy()

if df_lodd.empty:
    st.warning("Fant ingen lodd-rader der kategori/salgssted inneholder 'Lodd bilde'.")
    st.stop()

df_lodd["Bilde"] = df_lodd["LotteriKategori"].apply(extract_bilde)

df_lodd = df_lodd[df_lodd["Bilde"].notna()].copy()

if df_lodd.empty:
    st.error("Fant lodd-rader, men klarte ikke å lese bildebokstav fra 'LotteriKategori'.")
    st.stop()

# Brutto numeric
df_lodd["Brutto"] = pd.to_numeric(df_lodd["Brutto"], errors="coerce").fillna(0)

# Bygg Navn-kolonne (må finnes før groupby)
df_lodd["Navn"] = df_lodd.apply(build_full_name, axis=1)

if name_mode == "Kun fornavn":
    df_lodd["Navn"] = df_lodd["Navn"].apply(
        lambda s: str(s).strip().split(" ")[0] if str(s).strip() else "Ukjent"
    )

# Summer brutto per bilde/person FØRST (viktig for kjøp i flere omganger)
agg = (
    df_lodd.groupby(["Bilde", "Navn"], as_index=False)
    .agg(Brutto=("Brutto", "sum"))
)

# Regn ut lodd basert på total brutto
agg["Lodd_raw"] = agg["Brutto"] / float(loddpris)
if round_down:
    agg["Lodd"] = agg["Lodd_raw"].apply(lambda v: int(math.floor(v)))
else:
    agg["Lodd"] = agg["Lodd_raw"].round().astype(int)

agg["Lodd"] = agg["Lodd"].astype(int)

non_multiple = agg[(agg["Brutto"] % float(loddpris)) != 0]
if len(non_multiple) > 0:
    st.warning(
        f"{len(non_multiple)} kjøpere har totalbeløp som ikke går opp i loddpris ({loddpris} kr). "
        f"Appen {'runder ned' if round_down else 'runder'} til heltall lodd."
    )

    non_multiple = non_multiple[["Bilde", "Navn", "Brutto", "Lodd_raw"]].rename(columns={
        "Brutto": "Betalt sum",
        "Lodd_raw": "Betalt / Pris per lodd",
    })
    
    with st.expander("Se detaljer under:"):
        st.dataframe(non_multiple.sort_values(["Bilde","Navn"]), use_container_width=True, hide_index=True)
    
bilder = sorted(agg["Bilde"].unique().tolist())

st.success(f"✅ Fant {len(df_lodd)} lodd fordelt på {len(bilder)} bilder.")

st.subheader("Velg bilde under for å se deltakerliste og statistikk:")

if not bilder:
    st.error("Fant ingen bilder å vise. Sjekk kolonnen 'LotteriKategori' og regexen i extract_bilde().")
    st.stop()
tabs = st.tabs([f"Bilde {b}" for b in bilder])

WHEEL_URL = "https://wheelofnames.com/"

for tab, bilde in zip(tabs, bilder):
    with tab:
        left, right = st.columns([1.1, 0.9], gap="large")

        sub = agg[agg["Bilde"] == bilde].copy().sort_values(["Lodd", "Navn"], ascending=[False, True])

        # Clamp negative totals per person to 0 (refunds may net out)
        sub["Lodd_clamped"] = sub["Lodd"].apply(lambda x: max(int(x), 0))

        total_lodd = int(sub["Lodd_clamped"].sum())
        buyers = int((sub["Lodd_clamped"] > 0).sum())
        total_brutto = float(sub["Brutto"].sum())

        # Winner list text
        wheel_names = []
        for _, r in sub.iterrows():
            count = int(r["Lodd_clamped"])
            if count <= 0:
                continue
            wheel_names.extend([str(r["Navn"])] * count)

        wheel_text = "\n".join(wheel_names)

        with left:
            st.subheader("Trekning")
            st.caption("Kopiér listen og lim inn på Wheel of Names.")

            if wheel_text.strip():
                with st.expander("Vis liste (kopier med knappen øverst til høyre)", expanded=True):
                    st.code(wheel_text, language=None)
            else:
                st.info("Ingen lodd å kopiere for dette bildet (netto 0).")

            st.link_button("🎡 Åpne Wheel of Names", WHEEL_URL, use_container_width=True)

            st.divider()

            # Flyttet hit fra høyresiden:
            st.caption("Topp 10 kjøpere (etter antall lodd)")
            top10 = sub[sub["Lodd_clamped"] > 0].head(10).copy()

            if top10.empty:
                st.info("Ingen kjøpere med lodd > 0.")
            else:
                top10_view = (
                    top10[["Navn", "Lodd_clamped", "Brutto"]]
                    .rename(columns={
                        "Navn": "Navn",
                        "Lodd_clamped": "Lodd",
                        "Brutto": "Betalt sum",
                    })
                )
                st.dataframe(top10_view, use_container_width=True, hide_index=True)

        with right:
            k1, k2, k3 = st.columns(3)
            k1.metric("Kjøpere", buyers)
            k2.metric("Lodd", total_lodd)
            k3.metric("Total sum (kr)", f"{total_brutto:,.0f}".replace(",", " "))

            # Diagrammet blir igjen på høyre side
            if not top10.empty:
                chart_type = st.radio(
                    "Diagram", ["Stolpediagram", "Kakediagram"],
                    horizontal=True, key=f"chart_{bilde}"
                )

                if chart_type == "Stolpediagram":
                    fig = px.bar(top10, x="Navn", y="Lodd_clamped")
                    fig.update_layout(xaxis_title="", yaxis_title="Lodd")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    fig = px.pie(top10, names="Navn", values="Lodd_clamped")
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Ingen data å vise i diagrammet.")

                with st.expander(f"Vis tolket datagrunnlag for bilde {bilde}"):
                    df_dbg = df_lodd[df_lodd["Bilde"] == bilde].copy()
                    df_dbg["Lodd_raw"] = df_dbg["Brutto"] / float(loddpris)

                    # Velg kolonner som finnes
                    cols = [c for c in ["Tidspunkt", "Salgsdato", "LotteriKategori", "Navn", "Brutto", "Lodd_raw", "Melding"] if c in df_dbg.columns]
                    df_dbg = df_dbg[cols].rename(columns={
                        "LotteriKategori": "Kjøp",
                        "Brutto": "Betalt sum",
                        "Lodd_raw": "Betalt / Pris per lodd",
                    })

                    st.dataframe(df_dbg, use_container_width=True, hide_index=True)

# -----------------------------
# Footer
# -----------------------------

st.divider()

def svg_to_data_uri(svg_path: str) -> str | None:
    try:
        svg_bytes = Path(svg_path).read_bytes()
    except FileNotFoundError:
        return None
    b64 = base64.b64encode(svg_bytes).decode("utf-8")
    return f"data:image/svg+xml;base64,{b64}"

VERSION = "v1.0"
LAST_UPDATED = "21.02.2026" 

mailto_subject = quote("Henvendelse: App for Kunstlotteri - NHO")
mailto_link = f"mailto:casperalexei@gmail.com?subject={mailto_subject}"

gavin_logo_uri = svg_to_data_uri("assets/Logo.svg")       # hvis du har

logo_html = f'<img src="{gavin_logo_uri}" style="height:38px;" />' if gavin_logo_uri else "<b>Gavin Consulting</b>"


footer_html = f"""
<style>
  .gc-footer {{
    font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
    background: #f3f4f6;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 14px 16px;
    margin-top: 10px;
  }}
  .gc-footer-row {{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:16px;
    flex-wrap:wrap;
  }}
  .gc-footer-left {{
    display:flex;
    align-items:center;
    gap:12px;
    flex-wrap:wrap;
  }}
  .gc-footer-right {{
    display:flex;
    align-items:center;
    gap:12px;
    flex-wrap:wrap;
    justify-content:flex-end;
  }}
  .gc-icon-link img {{
    height: 24px;
    opacity: 0.85;
    transition: all 0.15s ease-in-out;
  }}
  .gc-icon-link img:hover {{
    opacity: 1.0;
    transform: translateY(-1px);
  }}
  .gc-btn {{
    background:#335E99;
    color:white;
    border:none;
    border-radius:10px;
    padding:8px 12px;
    cursor:pointer;
    font-size:14px;
  }}
  .gc-muted {{
    color:#6b7280;
    font-size:12px;
  }}
</style>

<div class="gc-footer">
  <div class="gc-footer-row">
    <div class="gc-footer-left">
      {logo_html}
      <div class="gc-muted">Versjon: {VERSION} • Sist oppdatert: {LAST_UPDATED}</div>
    </div>

    <div class="gc-footer-right">
      <div style="font-size:14px;">Dersom du har spørsmål, ta kontakt ved å trykke på knappen.</div>

      <a href="{mailto_link}" style="text-decoration:none;">
        <button class="gc-btn">✉️ Kontakt</button>
      </a>
    </div>
  </div>
</div>
"""

components.html(footer_html, height=100)