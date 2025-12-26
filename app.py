import streamlit as st
import pandas as pd
import glob
from pathlib import Path
import os
import sys

st.set_page_config(layout="wide")
st.title("ОС ⇄ IT merge по инвентарному номеру")

# =====================================================
# ENV CHECK: can we use tkinter?
# =====================================================
USE_TK = False
try:
    import tkinter as tk
    from tkinter import filedialog
    USE_TK = True
except Exception:
    USE_TK = False


def pick_folder():
    if not USE_TK:
        return None
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory()
    root.destroy()
    return folder


def norm_key(x):
    if pd.isna(x):
        return None
    return str(x).strip()


# =====================================================
# STEP 1 — SOURCE FILES
# =====================================================
st.subheader("1) Источник IT-данных")

it_files = []

if USE_TK:
    st.caption("Режим: локальный (выбор папки)")
    if "source_dir" not in st.session_state:
        st.session_state.source_dir = ""

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("📂 Выбрать папку"):
            selected = pick_folder()
            if selected:
                st.session_state.source_dir = selected

    with col2:
        source_dir = st.text_input(
            "Папка с IT-файлами (.xlsx)",
            value=st.session_state.source_dir,
            disabled=True
        )

    if not source_dir or not os.path.isdir(source_dir):
        st.warning("Выбери существующую папку")
        st.stop()

    it_files = [
        p for p in glob.glob(os.path.join(source_dir, "*.xlsx"))
        if not Path(p).name.startswith("~$")
    ]

else:
    st.caption("Режим: Cloud / Linux (загрузка файлов)")
    uploaded_files = st.file_uploader(
        "Загрузи IT-файлы (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.stop()

    it_files = uploaded_files


st.write(f"IT-файлов: **{len(it_files)}**")
if not it_files:
    st.error("Нет файлов для обработки")
    st.stop()


# =====================================================
# STEP 2 — SCAN UNIQUE COLUMNS
# =====================================================
st.subheader("2) Проверка уникальных колонок")

unique_cols = set()

for f in it_files:
    try:
        if USE_TK:
            df0 = pd.read_excel(f, nrows=0)
        else:
            df0 = pd.read_excel(f, nrows=0)
        unique_cols.update([str(c).strip() for c in df0.columns])
    except Exception:
        pass

unique_cols = sorted([c for c in unique_cols if c and c.lower() != "nan"])

st.write(f"Уникальных колонок: **{len(unique_cols)}**")
st.dataframe(pd.DataFrame({"column": unique_cols}), height=300)


# =====================================================
# STEP 3 — TARGET FILE
# =====================================================
st.subheader("3) Целевая ведомость ОС")

target_file = st.file_uploader("Ведомость ОС (.xlsx)", type=["xlsx"])
if not target_file:
    st.stop()

base = pd.read_excel(target_file)
base_cols = [str(c).strip() for c in base.columns]

st.write(f"Колонок в целевом файле: **{len(base_cols)}**")
st.dataframe(pd.DataFrame({"base_columns": base_cols}), height=240)


# =====================================================
# STEP 4 — MERGE SETTINGS
# =====================================================
st.subheader("4) Настройки объединения")

colA, colB = st.columns(2)
with colA:
    base_key = st.selectbox("Ключ в целевом файле", options=base_cols)
with colB:
    it_key = st.selectbox("Ключ в IT-файлах", options=unique_cols)

add_cols = st.multiselect(
    "Колонки для добавления в конец (любое количество)",
    options=[c for c in unique_cols if c != it_key],
    default=[]
)

if not add_cols:
    st.warning("Выбери хотя бы одну колонку")
    st.stop()


# =====================================================
# STEP 5 — RUN
# =====================================================
st.subheader("5) Выполнить объединение")

if st.button("MATCH", type="primary"):

    base = pd.read_excel(target_file)
    base = base.rename(columns={base_key: "inv_key"})
    base["inv_key"] = base["inv_key"].apply(norm_key)
    base_keys = set(base["inv_key"].dropna())

    it_frames = []
    unmatched_frames = []

    for f in it_files:
        try:
            df = pd.read_excel(f)
        except Exception as e:
            st.warning(f"Не прочитал файл: {e}")
            continue

        if it_key not in df.columns:
            continue

        df = df.rename(columns={it_key: "inv_key"})
        df["inv_key"] = df["inv_key"].apply(norm_key)

        um = df[~df["inv_key"].isin(base_keys)].copy()
        if not um.empty:
            um["Источник"] = Path(getattr(f, "name", f)).name
            unmatched_frames.append(um)

        existing = [c for c in add_cols if c in df.columns]
        if not existing:
            continue

        m = df[df["inv_key"].isin(base_keys)][["inv_key"] + existing]
        it_frames.append(m)

    if not it_frames:
        st.error("Нет совпадений по ключу")
        st.stop()

    it_all = pd.concat(it_frames, ignore_index=True)
    it_all = (
        it_all
        .groupby("inv_key", as_index=False)
        .agg(lambda s: s.dropna().iloc[0] if not s.dropna().empty else None)
    )

    result = base.merge(it_all, on="inv_key", how="left")
    unmatched_df = pd.concat(unmatched_frames, ignore_index=True) if unmatched_frames else pd.DataFrame()

    out_name = "os_merge_result.xlsx"
    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="MATCHED", index=False)
        unmatched_df.to_excel(writer, sheet_name="UNMATCHED", index=False)

    st.success("Готово")

    st.dataframe(result.head(200), height=300)

    with open(out_name, "rb") as f:
        st.download_button(
            "Скачать Excel",
            data=f,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
