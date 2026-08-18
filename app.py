import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from io import BytesIO

st.set_page_config(page_title="Сокращенный анализ локации", layout="wide")

st.title("🚗 ИИ-Агент: Сокращенный анализ локации")
st.write("Сравнение таблиц на листах 'АФ сокр' из двух отчетов с умной бизнес-подсветкой доходов и расходов.")

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

def is_colored(cell):
    """Проверяет, есть ли у ячейки цветная заливка"""
    if cell and cell.fill and cell.fill.fill_type:
        color = cell.fill.start_color.index
        if color and str(color) not in ['00000000', '0', 'FFFFFFFF', 'System_Color_Window']:
            return True
    return False

def get_colored_rows(file_bytes, sheet_name, header_idx, target_col_name):
    """Быстро находит строки с цветовой заливкой с защитой от пустых ячеек"""
    colored_rows = set()
    try:
        wb = openpyxl.load_workbook(file_bytes, data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            target_col_idx = None
            
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=header_idx, column=col).value
                if cell_val is not None and str(cell_val).strip() == target_col_name:
                    target_col_idx = col
                    break
                    
            if target_col_idx:
                for row_idx in range(header_idx + 1, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=target_col_idx)
                    if cell and cell.value is not None and is_colored(cell):
                        colored_rows.add(row_idx - header_idx - 1)
    except:
        pass
    return colored_rows

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

def make_color_excel(df, numeric_cols_list, type_col_name):
    """Отдельная функция для генерации разукрашенного Excel во избежание SyntaxError"""
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)
    towrite.seek(0)
    
    wb_export = openpyxl.load_workbook(towrite)
    ws_export = wb_export.active
    
    excel_fill_green = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    excel_fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    
    # Находим, в каком столбце Excel лежит маркер "Доходы Расходы"
    type_col_idx = 2
    for col in range(1, ws_export.max_column + 1):
        if str(ws_export.cell(row=1, column=col).value).strip() == type_col_name:
            type_col_idx = col
            break
            
    for r_idx in range(2, ws_export.max_row + 1):
        type_val_cell = ws_export.cell(row=r_idx, column=type_col_idx).value
        t_str = str(type_val_cell).strip().lower()
        row_is_income = "1" in t_str or "доход" in t_str
        
        for c_idx in range(1, ws_export.max_column + 1):
            col_name = str(ws_export.cell(row=1, column=c_idx).value).strip()
            if col_name in numeric_cols_list:
                cell_obj = ws_export.cell(row=r_idx, column=c_idx)
                cell_text = str(cell_obj.value)
                
                if "(+" in cell_text:
                    cell_obj.fill = excel_fill_green if row_is_income else excel_fill_red
                elif "(-" in cell_text:
                    cell_obj.fill = excel_fill_red if row_is_income else excel_fill_green
                    
    final_output = BytesIO()
    wb_export.save(final_output)
    final_output.seek(0)
    return final_output

# Основная логика приложения
if file_1 and file_2:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    pandas_header_index = int(header_row) - 1
    sheet_target = "АФ сокр"
    
    old_bytes = file_1.read()
    new_bytes = file_2.read()
    
    xl_1 = pd.ExcelFile(BytesIO(old_bytes))
    xl_2 = pd.ExcelFile(BytesIO(new_bytes))
    
    if sheet_target not in xl_1.sheet_names or sheet_target not in xl_2.sheet_names:
        st.error(f"❌ Ошибка: Лист '{sheet_target}' не найден в одном или обоих файлах!")
    else:
        df_1 = pd.read_excel(BytesIO(old_bytes), sheet_name=sheet_target, header=pandas_header_index)
        df_2 = pd.read_excel(BytesIO(new_bytes), sheet_name=sheet_target, header=pandas_header_index)
        
        df_1.columns = [str(c).strip() for c in df_1.columns]
        df_2.columns = [str(c).strip() for c in df_2.columns]
        
        if target_column not in df_1.columns or target_column not in df_2.columns:
            st.error(f"❌ Столбец '{target_column}' не найден на листе '{sheet_target}'. Проверьте настройки в боковом меню.")
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
            
            st.write("---")
            st.subheader("📥 Экспорт результатов в цветной Excel")
            st.write("Скачайте готовый отчет. Агент автоматически раскрасит ячейки доходов и расходов внутри файла Excel.")
            
            # Безопасный вызов изолированной функции генерации Excel
            excel_file_data = make_color_excel(df_result, numeric_cols, type_column)
            
            st.download_button(
                label="🟢 Скачать цветной отчет локации (Excel)",
                data=excel_file_data,
                file_name="Location_Analysis_Color_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для глубокого факторного анализа.")
