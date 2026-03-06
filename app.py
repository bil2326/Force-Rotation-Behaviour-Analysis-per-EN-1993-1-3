import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import StringIO, BytesIO

st.set_page_config(page_title="Analyse Courbe F–δ", layout="wide", page_icon="📊")

st.markdown("""
    <h2 style='color:#1e3a5f;'>📊 Analyse Courbe F–δ — Figure A.9 (EN 1993-1-3)</h2>
    <p style='color:#475569;'>Détermination automatique de δ_el et δ_pl par interpolation linéaire</p>
    <hr>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ Paramètres d'essai")
    s = st.number_input("Portée s (mm)", value=3600.0, step=1.0)
    e = st.number_input("Distance appui e (mm)", value=800.0, step=1.0)

    st.markdown("---")
    st.markdown("""
**Formule (Éq. A.4a) :**

`θ = 2(δ_pl − δ_e − δ_el) / (0.5·s − e)`

avec :

`δ_e = (δ_e_montante + δ_e_descendante) / 2`

- **δ_el** → branche montante F–δ
- **δ_pl** → branche descendante F–δ
- **δ_e** → moyenne des deux branches F–δ_e
""")

# ════════════════════════════════════════════════════════════════════════════
# FONCTIONS
# ════════════════════════════════════════════════════════════════════════════

def parse_paste(text):
    """Parse un collage Excel 2 colonnes → DataFrame [F, val]."""
    try:
        df = pd.read_csv(StringIO(text.strip()), sep=r"\t|;", engine="python",
                         header=None, on_bad_lines="skip")
        df = df.iloc[:, :2].copy()
        df.columns = ["F", "val"]
        df["F"]   = pd.to_numeric(df["F"].astype(str).str.replace(",", "."), errors="coerce")
        df["val"] = pd.to_numeric(df["val"].astype(str).str.replace(",", "."), errors="coerce")
        df = df.dropna().reset_index(drop=True)
        return df if len(df) >= 2 else None
    except:
        return None


def interp_branch(df, F_col, val_col, F_target):
    """
    Interpolation linéaire robuste sur UNE branche.
    - Regroupe par F (moyenne) pour éliminer les doublons
    - Trie par F croissant (requis par np.interp)
    - Retourne None si F_target hors plage
    """
    tmp = df[[F_col, val_col]].copy()
    tmp.columns = ["F", "v"]
    tmp = tmp.groupby("F", as_index=False)["v"].mean()
    tmp = tmp.sort_values("F").reset_index(drop=True)

    f = tmp["F"].values
    v = tmp["v"].values

    if f.min() <= F_target <= f.max():
        return float(np.interp(F_target, f, v))
    return None


def split_branches(df, f_col, val_col):
    """
    Sépare un DataFrame en branche montante et descendante.
    Retourne (fmax_idx, ascending, descending).
    """
    fmax_idx   = int(df[f_col].idxmax())
    ascending  = df.iloc[:fmax_idx + 1].reset_index(drop=True)
    descending = df.iloc[fmax_idx:].reset_index(drop=True)
    return fmax_idx, ascending, descending


# ════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ════════════════════════════════════════════════════════════════════════════
col1, col2 = st.columns([1, 1.6])

# ════════════════════════════════════════════════════════════════════════════
# COLONNE GAUCHE — saisies
# ════════════════════════════════════════════════════════════════════════════
with col1:

    # ── 1. Courbe F–δ ────────────────────────────────────────────────────────
    st.subheader("📋 Courbe F–δ (flèche milieu)")
    st.caption("Collez vos 2 colonnes **(F | δ)** depuis Excel :")

    paste_fdelta = st.text_area(
        "fdelta", height=200,
        placeholder="F (N)\tδ (mm)\n0\t0\n500\t0.55\n1000\t1.18\n...",
        label_visibility="collapsed", key="paste_fdelta"
    )

    df_points  = None
    fmax_idx   = None
    fmax_val   = None
    ascending  = None
    descending = None

    if paste_fdelta.strip():
        df_raw = parse_paste(paste_fdelta)
        if df_raw is not None:
            df_raw.columns = ["F", "delta"]
            df_points = df_raw

            fmax_idx, ascending, descending = split_branches(df_points, "F", "delta")
            fmax_val = float(df_points["F"].iloc[fmax_idx])

            def blabel(i):
                if i < fmax_idx:  return "↗ montante"
                if i == fmax_idx: return "⭐ F_max"
                return "↘ descendante"

            st.success(
                f"✅ **{len(df_points)} points** — "
                f"F_max = **{fmax_val:.3f} kN** | "
                f"montante : {fmax_idx + 1} pts | "
                f"descendante : {len(df_points) - fmax_idx} pts"
            )
            with st.expander("👁 Aperçu F–δ"):
                dv = df_points.copy()
                dv["F (kN)"]  = dv["F"].round(4)
                dv["Branche"] = [blabel(i) for i in range(len(df_points))]
                dv = dv.drop(columns=["F"])
                dv.index += 1
                st.dataframe(
                    dv.rename(columns={"delta": "δ (mm)"}),
                    use_container_width=True, height=200
                )
        else:
            st.error("❌ Données invalides — vérifiez vos colonnes.")

    st.markdown("---")

    # ── 2. Courbe F–δ_e ──────────────────────────────────────────────────────
    st.subheader("📋 Courbe F–δ_e (flèche appuis)")

    mode_de = st.radio(
        "Mode δ_e :",
        ["Valeur constante", "Courbe mesurée (Excel)"],
        horizontal=True
    )

    df_de         = None
    de_cst        = 0.0
    ascending_de  = None
    descending_de = None
    fmax_idx_de   = None

    if mode_de == "Valeur constante":
        de_cst = st.number_input(
            "δ_e (mm) — valeur unique pour tous les F",
            value=0.0, step=0.001, format="%.4f"
        )
        st.info("δ_e = constante appliquée à tous les niveaux de charge.")

    else:
        st.caption("Collez vos 2 colonnes **(F | δ_e)** depuis Excel :")
        paste_de = st.text_area(
            "de_paste", height=180,
            placeholder="F (N)\tδ_e (mm)\n0\t0\n500\t0.01\n1000\t0.02\n...",
            label_visibility="collapsed", key="paste_de"
        )

        if paste_de.strip():
            df_de_raw = parse_paste(paste_de)
            if df_de_raw is not None:
                df_de_raw.columns = ["F", "de"]
                df_de = df_de_raw

                fmax_idx_de, ascending_de, descending_de = split_branches(df_de, "F", "de")

                st.success(
                    f"✅ **{len(df_de)} points** δ_e importés — "
                    f"F : {df_de['F'].min():.3f} → {df_de['F'].max():.3f} kN | "
                    f"montante : {fmax_idx_de + 1} pts | "
                    f"descendante : {len(df_de) - fmax_idx_de} pts"
                )
                with st.expander("👁 Aperçu F–δ_e"):
                    dv2 = df_de.copy()
                    dv2["F (kN)"] = dv2["F"].round(4)
                    dv2 = dv2.drop(columns=["F"])
                    dv2.index += 1
                    st.dataframe(
                        dv2.rename(columns={"de": "δ_e (mm)"}),
                        use_container_width=True, height=180
                    )
            else:
                st.error("❌ Données δ_e invalides.")

    st.markdown("---")

    # ── 3. Niveaux F ─────────────────────────────────────────────────────────
    st.subheader("🎯 Niveaux F à analyser")
    st.caption(f"Entrez les valeurs en **kN** (un F par ligne) :")

    f_input = st.text_area(
        "fvals", height=140,
        placeholder="10\n20\n30\n35\n39",
        label_visibility="collapsed", key="f_input_area"
    )

    bc1, bc2 = st.columns(2)
    with bc1:
        if st.button("⚡ Auto 10 niveaux", use_container_width=True):
            if df_points is not None:
                st.session_state["f_auto"] = "\n".join(
                    str(round(fmax_val * (i + 1) / 11, 3)) for i in range(10)
                )
    with bc2:
        if st.button("📐 10%…95% F_max", use_container_width=True):
            if df_points is not None:
                st.session_state["f_auto"] = "\n".join(
                    str(round(fmax_val * p / 100, 3))
                    for p in [10, 20, 30, 40, 50, 60, 70, 80, 90, 95]
                )

    if "f_auto" in st.session_state and not f_input.strip():
        f_input = st.session_state["f_auto"]
        st.text_area(
            "(auto)", value=f_input, height=120,
            label_visibility="collapsed", disabled=True
        )

    run = st.button("▶ Analyser", type="primary", use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# ANALYSE
# ════════════════════════════════════════════════════════════════════════════
results_df = None

if run:
    if df_points is None:
        st.error("❌ Importez d'abord la courbe F–δ.")
    elif not f_input.strip():
        st.error("❌ Entrez au moins un niveau de charge F.")
    else:
        raw_f  = f_input.replace("\t", "\n").replace(";", "\n").split("\n")
        f_list = []
        for v in raw_f:
            try:
                # Les niveaux F sont saisis en kN (déjà convertis)
                fv = float(v.strip().replace(",", "."))
                if 0 < fv < fmax_val:
                    f_list.append(fv)
            except:
                pass

        if not f_list:
            st.error(f"❌ Aucune valeur F valide (0 < F < {fmax_val:.3f} kN).")
        else:
            rows, errs = [], []

            for F in sorted(set(f_list)):

                # δ_el — branche MONTANTE de F–δ
                del_val = interp_branch(ascending,  "F", "delta", F)
                # δ_pl — branche DESCENDANTE de F–δ
                dpl_val = interp_branch(descending, "F", "delta", F)

                if del_val is None or dpl_val is None:
                    errs.append(round(F, 4))
                    continue

                # ── CORRECTION BUG 2 : δ_e = moyenne montante + descendante ─
                if mode_de == "Valeur constante":
                    de_val_el = de_cst
                    de_val_pl = de_cst
                    de_moy    = de_cst          # constante → moyenne = elle-même
                else:
                    if ascending_de is not None and descending_de is not None:
                        de_val_el = interp_branch(ascending_de,  "F", "de", F)
                        de_val_pl = interp_branch(descending_de, "F", "de", F)
                        if de_val_el is None:
                            st.warning(f"⚠️ F={F:.3f} kN hors plage δ_e montante → 0 utilisé.")
                            de_val_el = 0.0
                        if de_val_pl is None:
                            st.warning(f"⚠️ F={F:.3f} kN hors plage δ_e descendante → 0 utilisé.")
                            de_val_pl = 0.0
                        # ── MOYENNE des deux branches (EN 1993-1-3 Fig. A.9) ─
                        de_moy = (de_val_el + de_val_pl) / 2.0
                    else:
                        de_val_el = 0.0
                        de_val_pl = 0.0
                        de_moy    = 0.0

                # ── Formule A.4a corrigée : δ_e = moyenne ───────────────────
                denom = 0.5 * s - e
                if abs(denom) < 1e-12:
                    st.error("❌ Dénominateur nul : vérifiez s et e.")
                    break
                theta = (2.0 * (dpl_val - de_moy - del_val)) / denom

                rows.append({
                    "F (kN)":                  round(F, 4),
                    "F / F_max (%)":           round(F / fmax_val * 100, 1),
                    "δ_el (mm)":               round(del_val, 4),
                    "δ_pl (mm)":               round(dpl_val, 4),
                    "δ_e montante (mm)":       round(de_val_el, 4),
                    "δ_e descendante (mm)":    round(de_val_pl, 4),
                    "δ_e moy (mm)":            round(de_moy, 4),      # ← colonne ajoutée
                    "δ_pl − δ_el (mm)":        round(dpl_val - del_val, 4),
                    "θ (rad) [Éq. A.4a]":      round(theta, 6),
                })

            if rows:
                results_df = pd.DataFrame(rows)
                st.session_state["results_df"]  = results_df
                st.session_state["results_pts"] = [
                    {"F": r["F (kN)"], "del": r["δ_el (mm)"], "dpl": r["δ_pl (mm)"]}
                    for r in rows
                ]
            if errs:
                st.warning(f"⚠️ F hors plage ignorés (kN) : {errs}")

elif "results_df" in st.session_state:
    results_df = st.session_state["results_df"]

# ════════════════════════════════════════════════════════════════════════════
# COLONNE DROITE — graphiques + résultats
# ════════════════════════════════════════════════════════════════════════════
with col2:

    # ── Graphique F–δ ─────────────────────────────────────────────────────
    st.subheader("📈 Courbe F–δ")
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.set_facecolor("#f8fafc")
    fig.patch.set_facecolor("white")

    if df_points is not None:
        up   = df_points.iloc[:fmax_idx + 1]
        down = df_points.iloc[fmax_idx:]

        # F est déjà en kN après conversion → axe Y cohérent
        ax.plot(up["delta"],   up["F"],   color="#2563eb", lw=2.2, label="Branche montante")
        ax.plot(down["delta"], down["F"], color="#ea580c", lw=2.2, label="Branche descendante")
        ax.scatter(
            df_points["delta"].iloc[fmax_idx],
            fmax_val,
            color="#dc2626", zorder=5, s=70,
            label=f"F_max = {fmax_val:.3f} kN"
        )

        for r in st.session_state.get("results_pts", []):
            # r["F"] est déjà en kN
            ax.axhline(r["F"], color="#cbd5e1", lw=0.7, ls="--", zorder=1)
            ax.scatter(r["del"], r["F"], color="#2563eb", marker="^", s=80, zorder=6)
            ax.scatter(r["dpl"], r["F"], color="#ea580c", marker="^", s=80, zorder=6)

        p1 = mpatches.Patch(color="#2563eb", label="▲ δ_el")
        p2 = mpatches.Patch(color="#ea580c", label="▲ δ_pl")
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=handles[:3] + [p1, p2], fontsize=8)

    ax.set_xlabel("Flèche nette δ (mm)", fontsize=10)
    ax.set_ylabel("Charge F (kN)", fontsize=10)
    ax.grid(True, alpha=0.4, lw=0.6)
    ax.spines[["top", "right"]].set_visible(False)
    st.pyplot(fig, use_container_width=True)

    # ── Graphique F–δ_e ──────────────────────────────────────────────────
    if df_de is not None and fmax_idx_de is not None:
        with st.expander("📈 Courbe F–δ_e (flèche aux appuis)"):
            fig2, ax2 = plt.subplots(figsize=(7, 3))
            up_de   = df_de.iloc[:fmax_idx_de + 1]
            down_de = df_de.iloc[fmax_idx_de:]
            ax2.plot(up_de["de"],   up_de["F"],   color="#2563eb", lw=2, label="Montante")
            ax2.plot(down_de["de"], down_de["F"], color="#ea580c", lw=2, label="Descendante")
            ax2.set_xlabel("δ_e (mm)", fontsize=10)
            ax2.set_ylabel("F (kN)", fontsize=10)
            ax2.legend(fontsize=8)
            ax2.grid(True, alpha=0.4)
            ax2.spines[["top", "right"]].set_visible(False)
            ax2.set_facecolor("#f8fafc")
            fig2.patch.set_facecolor("white")
            st.pyplot(fig2, use_container_width=True)

    # ── Résultats ─────────────────────────────────────────────────────────
    if results_df is not None:
        st.subheader("📋 Résultats")

        fmt = {
            "F (kN)":                  "{:.4f}",
            "F / F_max (%)":           "{:.1f}",
            "δ_el (mm)":               "{:.4f}",
            "δ_pl (mm)":               "{:.4f}",
            "δ_e montante (mm)":       "{:.4f}",
            "δ_e descendante (mm)":    "{:.4f}",
            "δ_e moy (mm)":            "{:.4f}",
            "δ_pl − δ_el (mm)":        "{:.4f}",
            "θ (rad) [Éq. A.4a]":      "{:.6f}",
        }

        styled = (results_df.style
            .applymap(lambda _: "color:#2563eb; font-weight:bold", subset=["δ_el (mm)"])
            .applymap(lambda _: "color:#ea580c; font-weight:bold", subset=["δ_pl (mm)"])
            .applymap(lambda _: "color:#7c3aed; font-weight:bold",
                      subset=["δ_e montante (mm)", "δ_e descendante (mm)", "δ_e moy (mm)"])
            .format(fmt)
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

        # ── Vérification manuelle affichée ────────────────────────────────
        with st.expander("🔍 Détail calcul θ (vérification)"):
            for _, row in results_df.iterrows():
                st.markdown(
                    f"**F = {row['F (kN)']:.4f} kN** → "
                    f"θ = 2×({row['δ_pl (mm)']:.4f} − {row['δ_e moy (mm)']:.4f} − {row['δ_el (mm)']:.4f}) "
                    f"/ (0.5×{s:.0f} − {e:.0f}) = **{row['θ (rad) [Éq. A.4a]']:.6f} rad**"
                )

        # ── Export ────────────────────────────────────────────────────────
        ec1, ec2 = st.columns(2)
        with ec1:
            csv_b = results_df.to_csv(
                index=False, sep=";", decimal=","
            ).encode("utf-8-sig")
            st.download_button(
                "⬇ CSV (Excel FR)", csv_b,
                "resultats.csv", "text/csv",
                use_container_width=True
            )
        with ec2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                results_df.to_excel(w, index=False, sheet_name="Résultats")
            st.download_button(
                "⬇ Excel (.xlsx)", buf.getvalue(),
                "resultats.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )