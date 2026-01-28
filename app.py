import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO

# --- НАСТРОЙКИ КОЛОНОК В EXCEL (Можешь поменять цифры под свой файл) ---
# Указываем индексы колонок (0 - это колонка A, 1 - B, 2 - C и т.д.)
COL_QUESTION = 1  # Где лежит сам вопрос (например, колонка B)
COL_A = 2         # Вариант A (например, колонка C)
COL_B = 3         # Вариант B
COL_C = 4         # Вариант C
COL_D = 5         # Вариант D
COL_E = 6         # Вариант E (для вопросов 36-40)
COL_F = 7         # Вариант F (для вопросов 36-40)

# Для контекстных вопросов (31-35), если они есть в Excel
# Если в твоем Excel этих данных нет, скрипт просто оставит пустоту.
COL_CTX_1 = 8     # Например, q31_first
COL_CTX_2 = 9     # Например, q31_second

def generate_doc(df, template_path):
    doc = DocxTemplate(template_path)
    context = {}

    # Пробегаемся по вопросам с 1 по 40
    # i - это номер вопроса, row_idx - номер строки в Excel (начинаем с 0 или 1 в зависимости от шапки)
    # Предполагаем, что данные начинаются со 2-й строки (индекс 0 в pandas, если header=0)
    
    for i in range(1, 41): 
        row_idx = i - 1  # 1-й вопрос лежит в 1-й строке данных
        
        # Если вопросов в файле меньше 40, останавливаемся
        if row_idx >= len(df):
            break

        # Базовые теги: q1, q1A, q1B...
        q_key = f"q{i}"
        
        # Безопасное получение данных (если ячейка пустая, ставим пробел)
        try:
            context[q_key] = str(df.iloc[row_idx, COL_QUESTION])
            context[f"{q_key}A"] = str(df.iloc[row_idx, COL_A])
            context[f"{q_key}B"] = str(df.iloc[row_idx, COL_B])
            context[f"{q_key}C"] = str(df.iloc[row_idx, COL_C])
            context[f"{q_key}D"] = str(df.iloc[row_idx, COL_D])
        except IndexError:
            continue # Если колонок не хватает, пропускаем

        # --- Специфика 31-35 (Контекст/Сопоставление) ---
        if 31 <= i <= 35:
            # Пытаемся достать q31_first и q31_second, если такие колонки есть
            try:
                context[f"{q_key}_first"] = str(df.iloc[row_idx, COL_CTX_1])
                context[f"{q_key}_second"] = str(df.iloc[row_idx, COL_CTX_2])
            except IndexError:
                context[f"{q_key}_first"] = ""
                context[f"{q_key}_second"] = ""

        # --- Специфика 36-40 (Множественный выбор, есть E и F) ---
        if 36 <= i <= 40:
            try:
                context[f"{q_key}E"] = str(df.iloc[row_idx, COL_E])
                context[f"{q_key}F"] = str(df.iloc[row_idx, COL_F])
            except IndexError:
                context[f"{q_key}E"] = ""
                context[f"{q_key}F"] = ""

    # Очистка данных от 'nan' (пустых ячеек Excel)
    clean_context = {k: (v if v != 'nan' else '') for k, v in context.items()}
    
    doc.render(clean_context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.set_page_config(page_title="Генератор Пробников", page_icon="📝")
st.title("Сборщик пробников Today UBT")
st.markdown("Загрузи **Excel** (вопросы в строках) и получи готовый **DOCX**.")

# 1. Загрузка Excel
uploaded_excel = st.file_uploader("1. Закинь Excel с вопросами", type=['xlsx', 'xlsm'])

# 2. Файл шаблона (он должен лежать рядом, но можно и загрузить для теста)
# Мы ищем файл 'template.docx' в папке скрипта.
try:
    template_file = "template.docx"
    # Проверка на наличие файла шаблона
    with open(template_file, "rb") as f:
        pass
except FileNotFoundError:
    st.error("⚠️ Файл template.docx не найден! Загрузи его в репозиторий.")
    st.stop()

if uploaded_excel:
    if st.button("Сгенерировать Word"):
        with st.spinner("Магия Python... ⏳"):
            try:
                # Читаем Excel. header=None значит, что мы берем по индексам колонок (0,1,2...),
                # даже если шапки нет или она кривая. skiprows=1 пропускает заголовок таблицы.
                df = pd.read_excel(uploaded_excel, header=None, skiprows=1)
                
                # Генерация
                doc_io = generate_doc(df, template_file)
                
                st.success("Готово! ✅")
                
                # Кнопка скачивания
                st.download_button(
                    label="📥 Скачать готовый тест",
                    data=doc_io,
                    file_name="Probnik_Result.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
                st.warning("Проверь структуру Excel: должно быть 40 строк данных.")