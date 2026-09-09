import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Сокращенный анализ локации", layout="wide")

st.title("🚗 ИИ-Агент: Сокращенный анализ локации")
st.write("Сравнение таблиц из двух отчетов с умной бизнес-подсветкой доходов и расходов.")

# Панель настроек в боковой панели
with st.sidebar:
    st.header("⚙️ Настройки структуры")
    target_column = st.text_input("Название столбца со статьями:", value="Статья")
    type_column = st.text_input("Название столбца типа (Доходы/Расходы):", value="Доходы Расходы")
    st.caption("ℹ️ Агент автоматически найдет строку с заголовками на листе, сопоставит ячейки и подсветит отклонения.")

# Блок загрузки файлов
col1, col2 = st.columns(2)
with col1:
    file_1 = st.file_uploader("📂 Загрузите файл 1 (Прошлый период / База)", type=["xlsx"])
with col2:
    file_2 = st.file_uploader("📂 Загрузите файл 2 (Текущий период / Отчет)", type=["xlsx"])

def clean_to_float(val):
    """Всеядная функция для приведения ячеек к числу с плавающей точкой"""
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str == "" or val_str == "-":
        return 0.0
    try:
        val_str = val_str.replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(val_str)
    except:
        return 0.0

def find_header_row_and_sheets(old_bytes, new_bytes, target_col):
    """Умный поиск строки заголовков и выбор правильного листа"""
    xl_1 = pd.ExcelFile(BytesIO(old_bytes))
    xl_2 = pd.ExcelFile(BytesIO(new_bytes))
    
    clean_sheets_1 = {str(name).strip().lower(): name for name in xl_1.sheet_names}
    clean_sheets_2 = {str(name).strip().lower(): name for name in xl_2.sheet_names}
    
    sheet_1_name, sheet_2_name = None, None
    for possible_name in ["аф сокр", "новая форма расходов"]:
        if possible_name in clean_sheets_1: sheet_1_name = clean_sheets_1[possible_name]
        if possible_name in clean_sheets_2: sheet_2_name = clean_sheets_2[possible_name]
        
    if not sheet_1_name or not sheet_2_name:
        return None, None, None, xl_1.sheet_names, xl_2.sheet_names
        
    detected_header_idx = None
    for r in range(12):
        try:
            df_test = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_2_name, header=r, nrows=2)
            cleaned_cols = [str(c).strip() for c in df_test.columns]
            if target_col in cleaned_cols:
                detected_header_idx = r
                break
        except:
            pass
            
    return sheet_1_name, sheet_2_name, detected_header_idx, xl_1.sheet_names, xl_2.sheet_names

def generate_html_report(df, numeric_cols, target_col_name, type_col_name, color_df):
    """Изолированная генерация HTML без f-строк во избежание конфликта скобок"""
    html = "<html><head><meta charset='utf-8'><style>"
    html += "body { font-family: Arial, sans-serif; padding: 20px; color: #333; }"
    html += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; margin-top:0; }"
    html += "table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11px; }"
    html += "th { background: #1E3A8A; color: white; padding: 6px; text-align: left; }"
    html += "td { padding: 6px; border-bottom: 1px solid #E5E7EB; }"
    html += "</style></head><body>"
    html += "<div style='background: white;'>"
    html += "<h2 style='margin-bottom:15px;'>Сокращенный анализ локации (Бизнес-отчет)</h2>"
    html += "<table><tr>"
    html += "<th>" + str(target_col_name) + "</th><th style='text-align: center;'>Тип</th>"
    for col in numeric_cols:
        html += "<th>" + str(col) + "</th>"
    html += "</tr>"
    
    for idx, row in df.iterrows():
        bg_row = "#F9FAFB" if idx % 2 == 0 else "#FFFFFF"
        t_str = str(row[type_col_name]).strip().lower()
        type_label = "Доход" if ("1" in t_str or "доход" in t_str) else "Расход"
        
        html += "<tr style='background: " + str(bg_row) + ";'>"
        html += "<td><b>" + str(row[target_col_name]) + "</b></td>"
        html += "<td style='text-align: center; color: #6B7280;'>" + str(type_label) + "</td>"
        
        for col in numeric_cols:
            cell_text = str(row[col])
            cell_style_raw = color_df.at[idx, col]
            extra_style = " " + str(cell_style_raw) if cell_style_raw else ""
            cell_style = "text-align: right;" + str(extra_style)
            html += "<td style='" + str(cell_style) + "'>" + str(cell_text) + "</td>"
            
        html += "</tr>"
        
    html += "</table></div></body></html>"
    return html

# Основная логика приложения
if file_1 and file_2:
    st.success("Файлы успешно загружены! Начинаю умный поиск структуры...")
    
    old_bytes = file_1.read()
    new_bytes = file_2.read()
    
    sheet_1, sheet_2, header_idx, sheets_1_all, sheets_2_all = find_header_row_and_sheets(old_bytes, new_bytes, target_column)
    
    if not sheet_1 or not sheet_2:
        st.error("❌ Ошибка: Целевой лист ('АФ сокр' или 'Новая форма расходов') не найден в одном или обоих файлах!")
        with st.expander("🔍 Посмотреть названия вкладок в ваших файлах"):
            st.write("**Листы в Файле 1:**", sheets_1_all)
            st.write("**Листы в Файле 2:**", sheets_2_all)
    elif header_idx is None:
        st.error(f"❌ Столбец '{target_column}' не найден в первых 12 строках на листе '{sheet_2}'. Проверьте точное написание заголовка.")
    else:
        st.info(f"⚙️ Структура определена автоматически. Заголовки найдены на строке {header_idx + 1}. Запускаю расчеты...")
        
        df_1 = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_1, header=header_idx)
        df_2 = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_2, header=header_idx)
        
        df_1.columns = [str(c).strip() for c in df_1.columns]
        df_2.columns = [str(c).strip() for c in df_2.columns]
        
        if type_column not in df_2.columns:
            st.error(f"❌ Столбец типа '{type_column}' не найден в новом файле. Доступные столбцы: {list(df_2.columns)}")
        else:
            df_1 = df_1.dropna(subset=[target_column])
            df_2 = df_2.dropna(subset=[target_column])
            df_1[target_column] = df_1[target_column].astype(str).str.strip()
            df_2[target_column] = df_2[target_column].astype(str).str.strip()
            
            numeric_cols = [col for col in df_2.columns if col != target_column and col != type_column and col in df_1.columns and not str(col).startswith('Unnamed:')]
            
            df_result_raw = df_2[[target_column, type_column] + numeric_cols].copy()
            df_result = df_result_raw.copy()
            
            for col in numeric_cols:
                df_result[col] = df_result[col].astype(object)
                
            df_1_indexed = df_1.set_index(target_column)
            color_matrix = pd.DataFrame('', index=df_result.index, columns=df_result.columns)
            
            STYLE_GREEN = 'background-color: #D1FAE5; color: #065F46;' 
            STYLE_RED = 'background-color: #FEE2E2; color: #991B1B;'   
            
            for idx, row in df_result_raw.iterrows():
                statya = row[target_column]
                raw_type_str = str(row[type_column]).strip().lower()
                is_income = "1" in raw_type_str or "доход" in raw_type_str
                
                for col in numeric_cols:
                    val_2 = row[col]
                    try:
                        val_1 = df_1_indexed.loc[statya, col]
                        if isinstance(val_1, pd.Series):
                            val_1 = val_1.iloc
                    except KeyError:
                        val_1 = 0.0
                        
                    val_2_float = clean_to_float(val_2)
                    val_1_float = clean_to_float(val_1)
                    delta = val_2_float - val_1_float
                    
                    if delta > 0:
                        df_result.at[idx, col] = f"{val_2_float:,.2f} (+{delta:,.2f})"
                        color_matrix.at[idx, col] = STYLE_GREEN if is_income else STYLE_RED
                    elif delta < 0:
                        df_result.at[idx, col] = f"{val_2_float:,.2f} (-{abs(delta):,.2f})"
                        color_matrix.at[idx, col] = STYLE_RED if is_income else STYLE_GREEN
                    else:
                        df_result.at[idx, col] = f"{val_2_float:,.2f}"
            
            def style_cells(df):
                return color_matrix
            
            st.subheader("📊 Результаты сравнительного анализа локации")
            st.write("Цветовая индикация адаптирована под экономику ДЦ: рост доходов и падение расходов подсвечены **зеленым**, падение доходов и рост расходов — **красным**.")
            st.dataframe(df_result.style.apply(style_cells, axis=None), use_container_width=True)
            
            st.write("---")
            st.subheader("🖨️ Печать и экспорт в PDF")
            st.write("Нажмите комбинацию клавиш **Ctrl + P** (или **Cmd + P** на Mac) прямо на этой странице браузера, чтобы сохранить этот отчет в PDF.")
            
            # Вызов безопасной изолированной функции HTML
            html_preview_data = generate_html_report(df_result, numeric_cols, target_column, type_column, color_matrix)
            st.components.v1.html(html_preview_data, height=500, scrolling=True)
            
            st.write("---")
            st.subheader("📥 Выгрузка в Excel (Стандарт)")
            
            towrite = BytesIO()
            df_result.to_excel(towrite, index=False, header=True)
            towrite.seek(0)
            
            st.download_button(
                label="🟢 Скачать итоговый анализ (Excel)",
                data=towrite,
                file_name="Location_Analysis_Report.xlsx",
