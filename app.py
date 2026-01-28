import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime  # <--- Добавили библиотеку для дат

# --- НАСТРОЙКИ КОЛОНОК ---
# (Оставляем как было в последней рабочей версии)
COL_QUESTION = 1  
COL_A = 2         
COL_B = 3         
COL_C = 4         
COL_D = 5         
COL_E = 6         
COL_F = 7         
COL_CTX_1 = 6     
COL_CTX_2 = 7     

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

        context[q_key] = get_val(COL_QUESTION)
        context[f"{q_key}A"] = get_val(COL_A)
        context[f"{q_key}B"] = get_val(COL_B)
        context[f"{q_key}C"] = get_val(COL_C)
        context[f"{q_key}D"] = get_val(COL_D)

        if 31 <= i <= 35:
            context[f"{q_key}_first"] = get_val(COL_CTX_1)
            context[f"{q_key}_second"] = get_val(COL_CTX_2)

        if 36 <= i <= 40:
            context[f"{q_key}E"] = get_val(COL_E)
            context[f"{q_key}F"] = get_val(COL_F)

    doc.render(context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- ИНТЕРФЕЙС ---
st.set_page_config(page_title="Генератор Пробников", page_icon="📅")
st.title("Генератор Пробников Today UBT")

uploaded_excel = st.file_uploader("Загрузите Excel", type=['xlsx', 'xlsm'])

try:
    template_file = "template.docx"
    with open(template_file, "rb") as f: pass
except FileNotFoundError:
    st.error("⚠️ Нет файла template.docx")
    st.stop()

if uploaded_excel:
    if st.button("Сгенерировать"):
        with st.spinner("Создаем файл..."):
            try:
                df = pd.read_excel(uploaded_excel, header=None, skiprows=1)
                doc_io = generate_doc(df, template_file)
                
                # --- ГЕНЕРАЦИЯ ИМЕНИ ФАЙЛА С ДАТОЙ ---
                # Получаем текущую дату в формате ДД_ММ_ГГГГ
                date_str = datetime.now().strftime("%d_%m_%Y")
                file_name_with_date = f"probnik_{date_str}.docx"
                
                st.success(f"Готово! Файл: {file_name_with_date}")
                
                st.download_button(
                    label=f"Скачать {file_name_with_date}",
                    data=doc_io,
                    file_name=file_name_with_date, # Вот сюда подставляем имя
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
