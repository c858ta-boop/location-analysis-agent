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
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=4)
    st.caption("ℹ️ Код 1 = Доходы (Рост=Зеленый, Падение=Красный). Код 2 = Расходы (Рост=Красный, Падение=Зеленый).")

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

# Основная логика приложения
if file_1 and file_2:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    pandas_header_index = int(header_row) - 1
    
    old_bytes = file_1.read()
    new_bytes = file_2.read()
    
    xl_1 = pd.ExcelFile(BytesIO(old_bytes))
    xl_2 = pd.ExcelFile(BytesIO(new_bytes))
    
    # Создаем карту очищенных имен листов
    clean_sheets_1 = {str(name).strip().lower(): name for name in xl_1.sheet_names}
    clean_sheets_2 = {str(name).strip().lower(): name for name in xl_2.sheet_names}
    
    sheet_1_real_name = None
    sheet_2_real_name = None
    
    for possible_name in ["аф сокр", "новая форма расходов"]:
        if possible_name in clean_sheets_1:
            sheet_1_real_name = clean_sheets_1[possible_name]
        if possible_name in clean_sheets_2:
            sheet_2_real_name = clean_sheets_2[possible_name]
            
    if not sheet_1_real_name or not sheet_2_real_name:
        st.error("❌ Ошибка: Целевой лист ('АФ сокр' или 'Новая форма расходов') не найден в одном или обоих файлах!")
        with st.expander("🔍 Посмотреть названия вкладок в ваших файлах"):
            st.write("**Листы в Файле 1:**", xl_1.sheet_names)
            st.write("**Листы в Файле 2:**", xl_2.sheet_names)
    else:
        df_1 = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_1_real_name, header=pandas_header_index)
        df_2 = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_2_real_name, header=pandas_header_index)
        
        df_1.columns = [str(c).strip() for c in df_1.columns]
        df_2.columns = [str(c).strip() for c in df_2.columns]
        
        if target_column not in df_1.columns or target_column not in df_2.columns:
            st.error(f"❌ Столбец '{target_column}' не найден на листе. Проверьте настройки в боковом меню.")
        elif type_column not in df_2.columns:
            st.error(f"❌ Столбец типа '{type_column}' не найден в новом файле. Проверьте заголовки.")
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
            
            # Построение вывода печатной формы прямо на экран (как в первом агенте, без try-except)
            st.write("---")
            st.subheader("🖨️ Печать и экспорт в PDF")
            st.write("Нажмите комбинацию клавиш **Ctrl + P** (или **Cmd + P** на Mac) прямо на этой странице браузера, чтобы сохранить этот отчет в PDF.")
            
            # Линейная генерация красивой HTML формы без двойных фигурных скобок CSS
            html_preview = "<html><head><meta charset='utf-8'><style>"
            html_preview += "body { font-family: Arial, sans-serif; margin: 20px; color: #333; }"
            html_preview += "h2 { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; font-size: 18px; margin-top:0; }"
            html_preview += "table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11px; }"
            html_preview += "th { background: #1E3A8A; color: white; padding: 6px; text-align: left; }"
            html_preview += "td { padding: 6px; border-bottom: 1px solid #E5E7EB; }"
            html_preview += "</style></head><body>"
            
            html_preview += "<div style='background: white;'>"
            html_preview += "<h2 style='margin-bottom:15px;'>Сокращенный анализ локации (Бизнес-отчет)</h2>"
            html_preview += "<table><tr>"
            html_preview += f"<th>{target_column}</th><th style='text-align: center;'>Тип</th>"
            for col in numeric_cols:
                html_preview += f"<th>{col}</th>"
            html_preview += "</tr>"
            
            for idx, row in df_result.iterrows():
                bg_row = "#F9FAFB" if idx % 2 == 0 else "#FFFFFF"
                t_str = str(row[type_column]).strip().lower()
                type_label = "Доход" if ("1" in t_str or "доход" in t_str) else "Расход"
                
                html_preview += f"<tr style='background: {bg_row};'>"
                html_preview += f"<td><b>{row[target_column]}</b></td>"
                html_preview += f"<td style='text-align: center; color: #6B7280;'>{type_label}</td>"
                
                for col in numeric_cols:
                    cell_text = str(row[col])
                    cell_style_raw = color_matrix.at[idx, col]
                    
                    # Безопасный плоский однострочник стилей
                    extra_style = " " + str(cell_style_raw) if cell_style_raw else ""
                    cell_style = f"text-align: right;{extra_style}"
                    html_preview += f"<td style='{cell_style}'>{cell_text}</td>"
                    
                html_preview += "</tr>"
                
            html_preview += "</table></div></body></html>"
            
            st.components.v1.html(html_preview, height=500, scrolling=True)
            
            st.write("---")
            st.subheader("📥 Выгрузка в Excel (Стандарт)")
            
            towrite = BytesIO()
            df_result.to_excel(towrite, index=False, header=True)
            towrite.seek(0)
            
            st.download_button(
                label="🟢 Скачать итоговый анализ (Excel)",
                data=towrite,
                file_name="Location_Analysis_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для глубокого факторного анализа.")
