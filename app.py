import streamlit as st
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from io import BytesIO

# --- НАСТРОЙКИ КОЛОНОК (ИСПРАВЛЕННЫЕ) ---
COL_QUESTION = 1  # Вопрос
COL_A = 2         # A
COL_B = 3         # B
COL_C = 4         # C
COL_D = 5         # D

# В твоем файле и Варианты E/F, и Подвопросы лежат в одних и тех же колонках (G и H)
# Индексы (считаем с 0): A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7
COL_E = 6         
COL_F = 7         
COL_CTX_1 = 6     # Первый подвопрос (тоже колонка G)
COL_CTX_2 = 7     # Второй подвопрос (тоже колонка H)

def generate_doc(df, template_path):
    doc = DocxTemplate(template_path)
    context = {}

    total_rows = len(df)
    
    for i in range(1, 41): 
        row_idx = i - 1
        
        if row_idx >= total_rows:
            break

        q_key = f"q{i}"
        
        def get_val(col_idx):
            try:
                val = df.iloc[row_idx, col_idx]
                return str(val).strip() if pd.notna(val) else ""
            except IndexError:
                return ""

        # Основные поля
        context[q_key] = get_val(COL_QUESTION)
        context[f"{q_key}A"] = get_val(COL_A)
        context[f"{q_key}B"] = get_val(COL_B)
        context[f"{q_key}C"] = get_val(COL_C)
        context[f"{q_key}D"] = get_val(COL_D)

        # Логика для 31-35 (Контекстные вопросы)
        # Теперь берем данные из колонок 6 и 7
        if 31 <= i <= 35:
            context[f"{q_key}_first"] = get_val(COL_CTX_1)
            context[f"{q_key}_second"] = get_val(COL_CTX_2)

        # Логика для 36-40 (Множественный выбор, варианты E и F)
        if 36 <= i <= 40:
            context[f"{q_key}E"] = get_val(COL_E)
            context[f"{q_key}F"] = get_val(COL_F)

    doc.render(context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- ИНТЕРФЕЙС ---
st.set_page_config(page_title="Генератор Пробников", page_icon="📝")
st.title("Генератор Пробников")

uploaded_excel = st.file_uploader("Загрузите Excel", type=['xlsx', 'xlsm'])

try:
    template_file = "template.docx"
    with open(template_file, "rb") as f: pass
except FileNotFoundError:
    st.error("⚠️ Нет файла template.docx")
    st.stop()

if uploaded_excel:
    if st.button("Сгенерировать"):
        with st.spinner("Работаем..."):
            try:
                # Читаем Excel (пропускаем 1-ю строку-шапку)
                df = pd.read_excel(uploaded_excel, header=None, skiprows=1)
                
                doc_io = generate_doc(df, template_file)
                
                st.success(f"Готово! Файл: {file_name_with_date}")
                st.download_button(
                    label=f"Скачать {file_name_with_date}",
                    data=doc_io,
                    file_name=file_name_with_date, # Вот сюда подставляем имя
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")

