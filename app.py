import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO

# --- НАСТРОЙКИ КОЛОНОК (Теперь идеально под твой файл) ---
COL_QUESTION = 1  # Вопрос (Колонка B)
COL_A = 2         # Вариант A (Колонка C)
COL_B = 3         # Вариант B (Колонка D)
COL_C = 4         # Вариант C (Колонка E)
COL_D = 5         # Вариант D (Колонка F)

# Для вопросов 36-40 (Множественный выбор), если они есть
COL_E = 6         # Вариант E (Колонка G)
COL_F = 7         # Вариант F (Колонка H)

# Для контекстных вопросов (31-35), если они есть
# Обычно они идут после основных вариантов. Если в твоем файле их нет в этих колонках,
# скрипт просто оставит пустоту, ошибок не будет.
COL_CTX_1 = 8     
COL_CTX_2 = 9     

def generate_doc(df, template_path):
    doc = DocxTemplate(template_path)
    context = {}

    # Проходим по строкам. 
    # Так как мы делаем skiprows=1 при чтении, то df.iloc[0] - это уже первый вопрос.
    # Нам нужно обработать ровно 40 вопросов (или сколько есть в файле).
    
    total_rows = len(df)
    
    for i in range(1, 41): 
        row_idx = i - 1  # 0-й индекс массива = 1-й вопрос теста
        
        # Защита: если в Excel меньше 40 строк, не ломаемся, а выходим
        if row_idx >= total_rows:
            break

        # Создаем ключи: q1, q1A, q1B...
        q_key = f"q{i}"
        
        # Функция для безопасного получения данных (чтобы не было 'nan')
        def get_val(col_idx):
            try:
                val = df.iloc[row_idx, col_idx]
                return str(val) if pd.notna(val) else ""
            except IndexError:
                return ""

        # Основное наполнение
        context[q_key] = get_val(COL_QUESTION)
        context[f"{q_key}A"] = get_val(COL_A)
        context[f"{q_key}B"] = get_val(COL_B)
        context[f"{q_key}C"] = get_val(COL_C)
        context[f"{q_key}D"] = get_val(COL_D)

        # Контекстные (31-35)
        if 31 <= i <= 35:
            context[f"{q_key}_first"] = get_val(COL_CTX_1)
            context[f"{q_key}_second"] = get_val(COL_CTX_2)

        # Множественный выбор (36-40)
        if 36 <= i <= 40:
            context[f"{q_key}E"] = get_val(COL_E)
            context[f"{q_key}F"] = get_val(COL_F)

    doc.render(context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- ИНТЕРФЕЙС ---
st.set_page_config(page_title="Генератор Пробников Today", page_icon="🎓")
st.title("Генератор Пробников (Excel -> Word)")

# Загрузка Excel
uploaded_excel = st.file_uploader("Загрузите файл Excel (заполненный!)", type=['xlsx', 'xlsm'])

# Проверка наличия шаблона
try:
    template_file = "template.docx"
    with open(template_file, "rb") as f:
        pass
except FileNotFoundError:
    st.error("⚠️ Не найден файл template.docx! Загрузи его в ту же папку.")
    st.stop()

if uploaded_excel:
    if st.button("Сгенерировать"):
        with st.spinner("Обработка..."):
            try:
                # Читаем Excel. 
                # skiprows=1 -> пропускаем шапку (Номер, Вопрос...), начинаем сразу с данных
                # header=None -> не используем имена колонок, работаем строго по номерам (0, 1, 2...)
                df = pd.read_excel(uploaded_excel, header=None, skiprows=1)
                
                doc_io = generate_doc(df, template_file)
                
                st.success("Файл готов!")
                st.download_button(
                    label="Скачать .docx",
                    data=doc_io,
                    file_name="Probnik_Final.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Произошла ошибка: {e}")
