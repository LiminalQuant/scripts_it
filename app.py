import streamlit as st
import pandas as pd
import glob
from pathlib import Path
import os

st.set_page_config(layout="wide")
st.title("ОС ⇄ IT merge по инвентарному номеру")

# ------------------- HELPERS -------------------
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

def safe_read_excel(path: str):
    return pd.read_excel(path)

# ------------------- STEP 1: SOURCE PATH + OUTPUT -------------------
st.subheader("1) Путь к папке с IT-файлами и место сохранения")
colA, colB, colC = st.columns([2, 2, 2])

with colA:
    source_dir = st.text_input("Папка с IT-файлами (.xlsx)", value="data/it")

with colB:
    out_dir = st.text_input("Папка для сохранения результата", value=".")

with colC:
    out_name = st.text_input("Имя итогового файла", value="os_merge_result.xlsx")

out_path = str(Path(out_dir) / out_name)

if not source_dir or not Path(source_dir).exists():
    st.warning("Укажи существующую папку с IT-файлами")
    st.stop()

it_files = list_xlsx_files(source_dir)
st.write(f"Найдено IT-файлов: **{len(it_files)}**")

if len(it_files) == 0:
    st.error("В папке нет .xlsx файлов (или только временные ~$/пустые)")
    st.stop()

# ------------------- STEP 2: UNIQUE COLUMNS SCAN -------------------
st.subheader("2) Проверка уникальных колонок (из IT-файлов)")

unique_cols = set()
cols_by_file = {}

with st.spinner("Сканирую заголовки..."):
    for f in it_files:
        cols = read_headers(f)
        cols_by_file[f] = cols
        unique_cols.update(cols)

unique_cols = sorted([c for c in unique_cols if c and c.lower() != "nan"])

col1, col2 = st.columns([2, 3])
with col1:
    st.write(f"Уникальных колонок найдено: **{len(unique_cols)}**")
with col2:
    st.caption("Это реальные заголовки колонок, как они есть в файлах. Дальше ты выберешь ключ и что добавить.")

st.dataframe(pd.DataFrame({"column": unique_cols}), height=320)

# ------------------- STEP 3: TARGET FILE UPLOAD -------------------
st.subheader("3) Добавь целевой файл (Ведомость ОС)")
target_file = st.file_uploader("Целевой Excel (.xlsx)", type=["xlsx"])

if not target_file:
    st.stop()

base = pd.read_excel(target_file)
base_cols = [str(c).strip() for c in base.columns]

st.write(f"Колонок в целевом файле: **{len(base_cols)}**")
st.dataframe(pd.DataFrame({"base_columns": base_cols}), height=240)

# ------------------- STEP 4: SELECT KEYS + COLUMNS -------------------
st.subheader("4) Выбор: по чему мэчим и что добавляем в конец")

colK1, colK2 = st.columns(2)
with colK1:
    base_key = st.selectbox("Ключ в целевом файле", options=base_cols)
with colK2:
    it_key = st.selectbox("Ключ в IT-файлах", options=unique_cols)

default_add = [c for c in unique_cols if c != it_key]
add_cols = st.multiselect(
    "Какие колонки добавить в конец (любое количество)",
    options=default_add,
    default=[]
)

if len(add_cols) == 0:
    st.warning("Выбери хотя бы одну колонку для добавления")
    st.stop()

# ------------------- STEP 5: RUN MERGE -------------------
st.subheader("5) Запуск объединения")

if st.button("MATCH", type="primary"):
    with st.spinner("Читаю целевой файл..."):
        base = pd.read_excel(target_file)
        base = base.rename(columns={base_key: "inv_key"})
        base["inv_key"] = base["inv_key"].apply(norm_key)
        base_keys = set(base["inv_key"].dropna().tolist())

    it_frames = []
    unmatched_frames = []
    matched_rows_total = 0

    with st.spinner("Читаю IT-файлы и собираю таблицу..."):
        for f in it_files:
            try:
                df = safe_read_excel(f)
            except Exception as e:
                st.warning(f"Не прочитал {Path(f).name}: {e}")
                continue

            if it_key not in df.columns:
                continue

            df = df.rename(columns={it_key: "inv_key"})
            df["inv_key"] = df["inv_key"].apply(norm_key)

            # UNMATCHED = строки, которых нет в базе по ключу
            um = df[~df["inv_key"].isin(base_keys)].copy()
            if not um.empty:
                um["Источник"] = Path(f).name
                unmatched_frames.append(um)

            # MATCHED = только совпавшие, только выбранные колонки (те, что реально есть в этом файле)
            existing = [c for c in add_cols if c in df.columns]
            if not existing:
                continue

            m = df[df["inv_key"].isin(base_keys)][["inv_key"] + existing].copy()
            if not m.empty:
                matched_rows_total += len(m)
                it_frames.append(m)

    if not it_frames:
        st.error("Совпадений по ключу нет (или выбранные колонки отсутствуют во всех файлах).")
        st.stop()

    # единая IT-таблица: по inv_key берём первое непустое значение по каждой колонке
    with st.spinner("Схлопываю IT-таблицу по inv_key (первое непустое)..."):
        it_all = pd.concat(it_frames, ignore_index=True)

        it_all = (
            it_all
            .groupby("inv_key", as_index=False)
            .agg(lambda s: s.dropna().iloc[0] if not s.dropna().empty else None)
        )

    # JOIN один раз в базу
    with st.spinner("Делаю JOIN в целевой файл..."):
        result = base.merge(it_all, on="inv_key", how="left")

    # UNMATCHED
    unmatched_df = pd.concat(unmatched_frames, ignore_index=True) if unmatched_frames else pd.DataFrame()

    # SAVE
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="MATCHED", index=False)
        unmatched_df.to_excel(writer, sheet_name="UNMATCHED", index=False)

    st.success(f"Готово. Сохранено: {out_path}")
    st.write(f"- Совпавших строк (до схлопывания): **{matched_rows_total}**")
    st.write(f"- Уникальных inv_key в IT после схлопывания: **{len(it_all)}**")
    st.write(f"- UNMATCHED строк: **{len(unmatched_df)}**")

    st.subheader("Превью MATCHED")
    st.dataframe(result.head(200), height=340)

    st.subheader("Превью UNMATCHED")
    st.dataframe(unmatched_df.head(200), height=340)

    with open(out_path, "rb") as f:
        st.download_button(
            "Скачать Excel (MATCHED + UNMATCHED)",
            data=f,
            file_name=Path(out_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
