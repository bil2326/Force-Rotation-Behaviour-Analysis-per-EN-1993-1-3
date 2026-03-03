import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import StringIO, BytesIO

st.set_page_config(page_title="Analyse Courbe F–δ", layout="wide", page_icon="📊")

st.markdown("""
    <h2 style='color:#1e3a5f;'>📊 Analyse Courbe F–δ — Figure A.9 (EN 1995)</h2>
    <p style='color:#475569;'>Détermination automatique de δ_el et δ_pl par interpolation linéaire</p>
    <hr>
""", unsafe_allow_html=True)

# ─── SIDEBAR ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Paramètres d'essai")
    s = st.number_input("Portée s (mm)", value=600.0, step=1.0)
    e = st.number_input("Distance appui e (mm)", value=50.0, step=1.0)
    st.markdown("---")
    st.markdown("""
**Formule (Éq. A.4a) :**

`θ = 2(δ_pl − δ_e − δ_el) / (0.5·s − e)`

- **δ_el** → branche montante  
- **δ_pl** → branche descendante  
- **δ_e**  → flèche aux appuis (interpolée pour chaque F)
""")

# ─── FONCTIONS ──────────────────────────────────────────────────────────────
def parse_paste(text):
    try:
        df = pd.read_csv(StringIO(text.strip()), sep=r"\t|;", engine="python",
                         header=None, on_bad_lines="skip")
        df = df.iloc[:, :2].copy()
        df.columns = ["F", "val"]
        df["F"]   = pd.to_numeric(df["F"].astype(str).str.replace(",","."), errors="coerce")
        df["val"] = pd.to_numeric(df["val"].astype(str).str.replace(",","."), errors="coerce")
        df = df.dropna().reset_index(drop=True)
        return df if len(df) >= 2 else None
    except:
        return None

def interp_branch(df, F_col, val_col, F_target):
    f = df[F_col].values
    v = df[val_col].values
    if f.min() <= F_target <= f.max():
        return float(np.interp(F_target, f, v))
    return None

# ─── COLONNES ───────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1.6])

with col1:

    # ── Courbe F–δ ──────────────────────────────────────────────────────────
    st.subheader("📋 Courbe F–δ (flèche milieu)")
    st.caption("Collez vos 2 colonnes **(F | δ)** depuis Excel :")
    paste_fdelta = st.text_area("fdelta", height=200,
        placeholder="F (N)\tδ (mm)\n0\t0\n500\t0.55\n1000\t1.18\n...",
        label_visibility="collapsed")

    df_points = None
    fmax_idx  = None

    if paste_fdelta.strip():
        df_raw = parse_paste(paste_fdelta)
        if df_raw is not None:
            df_raw.columns = ["F", "delta"]
            df_points = df_raw
            fmax_idx  = int(df_points["F"].idxmax())
            fmax_val  = df_points["F"].iloc[fmax_idx]

            def blabel(i):
                if i < fmax_idx:  return "↗ montante"
                if i == fmax_idx: return "⭐ F_max"
                return "↘ descendante"

            st.success(f"✅ **{len(df_points)} points** — F_max = **{fmax_val:.1f} N** "
                       f"| montante : {fmax_idx+1} pts | descendante : {len(df_points)-fmax_idx} pts")
            with st.expander("👁 Aperçu F–δ"):
                dv = df_points.copy()
                dv["Branche"] = [blabel(i) for i in range(len(df_points))]
                dv.index += 1
                st.dataframe(dv.rename(columns={"F":"F (N)","delta":"δ (mm)"}),
                             use_container_width=True, height=200)
        else:
            st.error("❌ Données invalides.")

    st.markdown("---")

    # ── Courbe F–δ_e ────────────────────────────────────────────────────────
    st.subheader("📋 Courbe F–δ_e (flèche appuis)")

    mode_de = st.radio("Mode δ_e :",
        ["Valeur constante", "Courbe mesurée (Excel)"], horizontal=True)

    df_de  = None
    de_cst = 0.0

    if mode_de == "Valeur constante":
        de_cst = st.number_input("δ_e (mm) — valeur unique",
                                 value=0.0, step=0.001, format="%.4f")
        st.info("δ_e = constante pour tous les niveaux de charge.")
    else:
        st.caption("Collez vos 2 colonnes **(F | δ_e)** depuis Excel :")
        paste_de = st.text_area("de_paste", height=180,
            placeholder="F (N)\tδ_e (mm)\n0\t0\n500\t0.01\n1000\t0.02\n...",
            label_visibility="collapsed", key="paste_de")

        if paste_de.strip():
            df_de_raw = parse_paste(paste_de)
            if df_de_raw is not None:
                df_de_raw.columns = ["F", "de"]
                df_de = df_de_raw
                st.success(f"✅ **{len(df_de)} points** δ_e importés "
                           f"(F : {df_de['F'].min():.0f} → {df_de['F'].max():.0f} N)")
                with st.expander("👁 Aperçu F–δ_e"):
                    dv2 = df_de.copy(); dv2.index += 1
                    st.dataframe(dv2.rename(columns={"F":"F (N)","de":"δ_e (mm)"}),
                                 use_container_width=True, height=180)
            else:
                st.error("❌ Données δ_e invalides.")

    st.markdown("---")

    # ── Niveaux F ───────────────────────────────────────────────────────────
    st.subheader("🎯 Niveaux F à analyser")
    st.caption("Collez une colonne Excel ou tapez (un F par ligne) :")
    f_input = st.text_area("fvals", height=140,
        placeholder="1000\n2000\n3000\n4000\n4500",
        label_visibility="collapsed")

    bc1, bc2 = st.columns(2)
    with bc1:
        if st.button("⚡ Auto 10 niveaux", use_container_width=True):
            if df_points is not None:
                fv = df_points["F"].iloc[fmax_idx]
                st.session_state["f_auto"] = "\n".join(str(round(fv*(i+1)/11)) for i in range(10))
    with bc2:
        if st.button("📐 10%…95% F_max", use_container_width=True):
            if df_points is not None:
                fv = df_points["F"].iloc[fmax_idx]
                st.session_state["f_auto"] = "\n".join(str(round(fv*p/100)) for p in [10,20,30,40,50,60,70,80,90,95])

    if "f_auto" in st.session_state and not f_input.strip():
        f_input = st.session_state["f_auto"]
        st.text_area("(auto)", value=f_input, height=120, label_visibility="collapsed", disabled=True)

    run = st.button("▶ Analyser", type="primary", use_container_width=True)

# ─── ANALYSE ────────────────────────────────────────────────────────────────
results_df = None

if run and df_points is not None and f_input.strip():
    fmax_val  = df_points["F"].iloc[fmax_idx]
    ascending = df_points.iloc[:fmax_idx+1].reset_index(drop=True)
    desc_inv  = df_points.iloc[fmax_idx:].iloc[::-1].reset_index(drop=True)

    raw_f  = f_input.replace("\t","\n").replace(";","\n").split("\n")
    f_list = []
    for v in raw_f:
        try:
            fv = float(v.strip().replace(",","."))
            if 0 < fv < fmax_val: f_list.append(fv)
        except: pass

    if not f_list:
        st.error(f"❌ Aucune valeur F valide (0 < F < {fmax_val:.1f} N).")
    else:
        rows, errs = [], []
        for F in sorted(set(f_list)):
            del_val = interp_branch(ascending, "F", "delta", F)
            dpl_val = interp_branch(desc_inv,  "F", "delta", F)

            if del_val is None or dpl_val is None:
                errs.append(F); continue

            # δ_e : constante ou interpolée
            if mode_de == "Valeur constante":
                de_val = de_cst
            else:
                de_val = interp_branch(df_de, "F", "de", F) if df_de is not None else 0.0
                if de_val is None:
                    st.warning(f"⚠️ F={F} N hors plage courbe δ_e → δ_e = 0 utilisé.")
                    de_val = 0.0

            theta = (2 * (dpl_val - de_val - del_val)) / (0.5 * s - e)
            rows.append({
                "F (N)":              round(F, 1),
                "F / F_max (%)":      round(F / fmax_val * 100, 1),
                "δ_el (mm)":          round(del_val, 4),
                "δ_pl (mm)":          round(dpl_val, 4),
                "δ_e (mm)":           round(de_val, 4),
                "δ_pl − δ_el (mm)":   round(dpl_val - del_val, 4),
                "θ (rad) [Éq. A.4a]": round(theta, 6),
            })

        if rows:
            results_df = pd.DataFrame(rows)
            st.session_state["results_df"]  = results_df
            st.session_state["results_pts"] = [
                {"F": r["F (N)"], "del": r["δ_el (mm)"], "dpl": r["δ_pl (mm)"]} for r in rows]
        if errs:
            st.warning(f"⚠️ Hors plage ignorés : {errs}")

elif "results_df" in st.session_state:
    results_df = st.session_state["results_df"]

# ─── COLONNE DROITE ─────────────────────────────────────────────────────────
with col2:
    st.subheader("📈 Courbe F–δ")
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.set_facecolor("#f8fafc"); fig.patch.set_facecolor("white")

    if df_points is not None:
        up   = df_points.iloc[:fmax_idx+1]
        down = df_points.iloc[fmax_idx:]
        ax.plot(up["delta"],   up["F"],   color="#2563eb", lw=2.2, label="Branche montante")
        ax.plot(down["delta"], down["F"], color="#ea580c", lw=2.2, label="Branche descendante")
        ax.scatter(df_points["delta"].iloc[fmax_idx], df_points["F"].iloc[fmax_idx],
                   color="#dc2626", zorder=5, s=70,
                   label=f"F_max = {df_points['F'].iloc[fmax_idx]:.0f} N")

        for r in st.session_state.get("results_pts", []):
            ax.axhline(r["F"], color="#cbd5e1", lw=0.7, ls="--", zorder=1)
            ax.scatter(r["del"], r["F"], color="#2563eb", marker="^", s=80, zorder=6)
            ax.scatter(r["dpl"], r["F"], color="#ea580c", marker="^", s=80, zorder=6)

        p1 = mpatches.Patch(color="#2563eb", label="▲ δ_el")
        p2 = mpatches.Patch(color="#ea580c", label="▲ δ_pl")
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=handles[:3]+[p1,p2], fontsize=8)

    ax.set_xlabel("Flèche nette δ (mm)", fontsize=10)
    ax.set_ylabel("Charge F (N)", fontsize=10)
    ax.grid(True, alpha=0.4, lw=0.6)
    ax.spines[["top","right"]].set_visible(False)
    st.pyplot(fig, use_container_width=True)

    # Graphique δ_e si courbe saisie
    if df_de is not None:
        with st.expander("📈 Courbe F–δ_e"):
            fig2, ax2 = plt.subplots(figsize=(7, 3))
            ax2.plot(df_de["de"], df_de["F"], color="#7c3aed", lw=2, marker="o", ms=4)
            ax2.set_xlabel("δ_e (mm)", fontsize=10)
            ax2.set_ylabel("F (N)", fontsize=10)
            ax2.grid(True, alpha=0.4)
            ax2.spines[["top","right"]].set_visible(False)
            ax2.set_facecolor("#f8fafc"); fig2.patch.set_facecolor("white")
            st.pyplot(fig2, use_container_width=True)

    # ── Résultats ────────────────────────────────────────────────────────────
    if results_df is not None:
        st.subheader("📋 Résultats")
        styled = (results_df.style
            .applymap(lambda _: "color:#2563eb;font-weight:bold", subset=["δ_el (mm)"])
            .applymap(lambda _: "color:#ea580c;font-weight:bold", subset=["δ_pl (mm)"])
            .applymap(lambda _: "color:#7c3aed;font-weight:bold", subset=["δ_e (mm)"])
            .format({
                "F (N)": "{:.1f}", "F / F_max (%)": "{:.1f}",
                "δ_el (mm)": "{:.4f}", "δ_pl (mm)": "{:.4f}", "δ_e (mm)": "{:.4f}",
                "δ_pl − δ_el (mm)": "{:.4f}", "θ (rad) [Éq. A.4a]": "{:.6f}",
            }))
        st.dataframe(styled, use_container_width=True, hide_index=True)

        ec1, ec2 = st.columns(2)
        with ec1:
            csv_b = results_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
            st.download_button("⬇ CSV", csv_b, "resultats.csv", "text/csv", use_container_width=True)
        with ec2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                results_df.to_excel(w, index=False, sheet_name="Résultats")
            st.download_button("⬇ Excel", buf.getvalue(), "resultats.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)