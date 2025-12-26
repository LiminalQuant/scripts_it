import streamlit as st
import pandas as pd
import glob
from pathlib import Path
import os

# ===== folder picker (Windows / local) =====
import tkinter as tk
from tkinter import filedialog


def pick_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory()
    root.destroy()
    return folder


# ================== STREAMLIT ==================
st.set_page_config(layout="wide")
st.title("ОС ⇄ IT merge по инвентарному номеру")


# ================== HELPERS ==================
def list_xlsx_files(folder: str):
    files = []
    for p in glob.glob(os.path.join(folder, "*.xlsx")):
        name = Path(p).name
        if name.startswith("~$"):
            continue
        files.append(p)
    return sorted(files)


def read_headers(file_path: str):
    try:
        df0 = pd.read_excel(file_path, nrows=0)
        return [str(c).strip() for c in df0.columns]
    except Exception:
        return []


def norm_key(x):
    if pd.isna(x):
        return None
    return str(x).strip()


# ================== STEP 1 ==================
st.subheader("1) Выбор папки с IT-файлами и места сохранения")

if "source_dir" not in st.session_state:
    st.session_state.source_dir = ""

colA, colB, colC = st.columns([2, 2, 2])

with colA:
    if st.button("📂 Выбрать папку с IT-файлами"):
        selected = pick_folder()
        if selected:
            st.session_state.source_dir = selected

with colB:
    out_dir = st.text_input("Папка для сохранения результата", value=".")

with colC:
    out_name = st.text_input("Имя итогового файла", value="os_merge_result.xlsx")

source_dir = st.text_input(
    "Папка с IT-файлами (.xlsx)",
    value=st.session_state.source_dir,
    disabled=True
)

out_path = str(Path(out_dir) / out_name)

if not source_dir or not os.path.isdir(source_dir):
    st.warning("Выбери существующую папку с IT-файлами")
    st.stop()

it_files = list_xlsx_files(source_dir)
st.write(f"Найдено IT-файлов: **{len(it_files)}**")

if not it_files:
    st.error("В выбранной папке нет .xlsx файлов")
    st.stop()


# ================== STEP 2 ==================
st.subheader("2) Проверка уникальных колонок (из IT-файлов)")

unique_cols = set()

with st.spinner("Сканирую заголовки..."):
    for f in it_files:
        cols = read_headers(f)
        unique_cols.update(cols)

unique_cols = sorted([c for c in unique_cols if c and c.lower() != "nan"])

st.write(f"Уникальных колонок найдено: **{len(unique_cols)}**")
st.dataframe(pd.DataFrame({"column": unique_cols}), height=320)


# ================== STEP 3 ==================
st.subheader("3) Добавь целевой файл (Ведомость ОС)")
target_file = st.file_uploader("Целевой Excel (.xlsx)", type=["xlsx"])

if not target_file:
    st.stop()

base = pd.read_excel(target_file)
base_cols = [str(c).strip() for c in base.columns]

st.write(f"Колонок в целевом файле: **{len(base_cols)}**")
st.dataframe(pd.DataFrame({"base_columns": base_cols}), height=240)


# ================== STEP 4 ==================
st.subheader("4) Выбор: по чему мэчим и что добавляем в конец")

colK1, colK2 = st.columns(2)

with colK1:
    base_key = st.selectbox("Ключ в целевом файле", options=base_cols)

with colK2:
    it_key = st.selectbox("Ключ в IT-файлах", options=unique_cols)

add_cols = st.multiselect(
    "Колонки, которые добавить в конец (любое количество)",
    options=[c for c in unique_cols if c != it_key],
    default=[]
)

if not add_cols:
    st.warning("Выбери хотя бы одну колонку для добавления")
    st.stop()


# ================== STEP 5 ==================
st.subheader("5) Запуск объединения")

if st.button("MATCH", type="primary"):

    # --- BASE ---
    base = pd.read_excel(target_file)
    base = base.rename(columns={base_key: "inv_key"})
    base["inv_key"] = base["inv_key"].apply(norm_key)
    base_keys = set(base["inv_key"].dropna())

    it_frames = []
    unmatched_frames = []

    # --- IT FILES ---
    with st.spinner("Читаю IT-файлы и собираю данные..."):
        for f in it_files:
            try:
                df = pd.read_excel(f)
            except Exception as e:
                st.warning(f"Не прочитал {Path(f).name}: {e}")
                continue

            if it_key not in df.columns:
                continue

            df = df.rename(columns={it_key: "inv_key"})
            df["inv_key"] = df["inv_key"].apply(norm_key)

            # UNMATCHED
            um = df[~df["inv_key"].isin(base_keys)].copy()
            if not um.empty:
                um["Источник"] = Path(f).name
                unmatched_frames.append(um)

            # MATCHED (берём только существующие колонки)
            existing = [c for c in add_cols if c in df.columns]
            if not existing:
                continue

            m = df[df["inv_key"].isin(base_keys)][["inv_key"] + existing]
            if not m.empty:
                it_frames.append(m)

    if not it_frames:
        st.error("Совпадений по ключу нет или выбранные колонки отсутствуют")
        st.stop()

    # --- COLLAPSE ---
    it_all = pd.concat(it_frames, ignore_index=True)
    it_all = (
        it_all
        .groupby("inv_key", as_index=False)
        .agg(lambda s: s.dropna().iloc[0] if not s.dropna().empty else None)
    )

    # --- MERGE ---
    result = base.merge(it_all, on="inv_key", how="left")

    unmatched_df = (
        pd.concat(unmatched_frames, ignore_index=True)
        if unmatched_frames else pd.DataFrame()
    )

    # --- SAVE ---
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="MATCHED", index=False)
        unmatched_df.to_excel(writer, sheet_name="UNMATCHED", index=False)

    st.success(f"Готово. Файл сохранён: {out_path}")

    st.subheader("Превью MATCHED")
    st.dataframe(result.head(200), height=350)

    st.subheader("Превью UNMATCHED")
    st.dataframe(unmatched_df.head(200), height=350)

    with open(out_path, "rb") as f:
        st.download_button(
            "Скачать Excel (MATCHED + UNMATCHED)",
            data=f,
            file_name=Path(out_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
