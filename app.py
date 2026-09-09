import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Сокращенный анализ локации", layout="wide")

st.title("🚗 ИИ-Агент: Сокращенный анализ локации")
st.write("Сравнение таблиц на листах 'АФ сокр' из двух отчетов с цветовой и знаковой индикацией изменений.")

# Боковая панель настроек
with st.sidebar:
    st.header("⚙️ Настройки структуры")
    target_column = st.text_input("Название столбца со статьями:", value="Статья")
    header_row = st.number_input("Строка с заголовками (в Excel нумерация с 1):", min_value=1, value=4)
    st.caption("ℹ️ Агент найдет столбец 'Статья' на листе 'АФ сокр', сопоставит все ячейки и подсветит отклонения.")

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

# Запуск основного интерфейса только при наличии обоих файлов
if file_1 and file_2:
    st.success("Файлы успешно загружены! Начинаю факторный анализ...")
    
    try:
        pandas_header_index = int(header_row) - 1
        sheet_target = "АФ сокр"
        
        # Клонируем файлы в независимые буферы памяти для стабильного чтения
        bytes_1 = file_1.read()
        bytes_2 = file_2.read()
        
        xl_1 = pd.ExcelFile(BytesIO(bytes_1))
        xl_2 = pd.ExcelFile(BytesIO(new_bytes_1)) if 'new_bytes_1' in locals() else pd.ExcelFile(BytesIO(bytes_2))
        
        if sheet_target not in xl_1.sheet_names or sheet_target not in xl_2.sheet_names:
            st.error(f"❌ Ошибка: Лист '{sheet_target}' не найден в одном или обоих файлах!")
        else:
            # Загружаем датафреймы
            df_1 = pd.read_excel(BytesIO(bytes_1), sheet_name=sheet_target, header=pandas_header_index)
            df_2 = pd.read_excel(BytesIO(bytes_2), sheet_name=sheet_target, header=pandas_header_index)
            
            # Очищаем заголовки от пробелов по краям
            df_1.columns = [str(c).strip() for c in df_1.columns]
            df_2.columns = [str(c).strip() for c in df_2.columns]
            
            if target_column not in df_1.columns or target_column not in df_2.columns:
                st.error(f"❌ Столбец '{target_column}' не найден на листе '{sheet_target}'. Проверьте настройки в боковом меню.")
            else:
                # Фильтруем пустые значения в столбце Статья и приводим к чистому строковому виду
                df_1 = df_1.dropna(subset=[target_column])
                df_2 = df_2.dropna(subset=[target_column])
                df_1[target_column] = df_1[target_column].astype(str).str.strip()
                df_2[target_column] = df_2[target_column].astype(str).str.strip()
                
                # Находим все числовые столбцы для сравнения
                numeric_cols = [col for col in df_2.columns if col != target_column and col in df_1.columns and not str(col).startswith('Unnamed:')]
                
                # Создаем каркас результирующей таблицы на основе Файла 2
                df_result_raw = df_2[[target_column] + numeric_cols].copy()
                df_result = df_result_raw.copy()
                
                for col in numeric_cols:
                    df_result[col] = df_result[col].astype(object)
                    
                df_1_indexed = df_1.set_index(target_column)
                
                # Матрица цветов для отображения в Streamlit
                color_matrix = pd.DataFrame('', index=df_result.index, columns=df_result.columns)
                
                # Константы стилей для ячеек
                STYLE_GREEN = 'background-color: #D1FAE5; color: #065F46;' 
                STYLE_RED = 'background-color: #FEE2E2; color: #991B1B;'   
                
                # Пересчитываем каждую ячейку
                for idx, row in df_result_raw.iterrows():
                    statya = row[target_column]
                    
                    for col in numeric_cols:
                        val_2 = row[col]
                        
                        try:
                            val_1 = df_1_indexed.loc[statya, col]
                            if isinstance(val_1, pd.Series):
                                val_1 = val_1.iloc[0]
                        except KeyError:
                            val_1 = 0.0
                            
                        val_2_float = clean_to_float(val_2)
                        val_1_float = clean_to_float(val_1)
                        
                        delta = val_2_float - val_1_float
                        
                        if delta > 0:
                            df_result.at[idx, col] = f"{val_2_float:,.2f} (+{delta:,.2f})"
                            color_matrix.at[idx, col] = STYLE_GREEN
                        elif delta < 0:
                            df_result.at[idx, col] = f"{val_2_float:,.2f} (-{abs(delta):,.2f})"
                            color_matrix.at[idx, col] = STYLE_RED
                        else:
                            df_result.at[idx, col] = f"{val_2_float:,.2f}"
                
                def style_cells(df):
                    return color_matrix
                
                st.subheader("📊 Результаты сравнительного анализа локации")
                st.write("В ячейках указано актуальное значение из Файла 2, а в скобках — разница к Файлу 1:")
                st.dataframe(df_result.style.apply(style_cells, axis=None), use_container_width=True)
                
                st.write("---")
                st.subheader("📥 Экспорт результатов")
                
                towrite = BytesIO()
                df_result.to_excel(towrite, index=False, header=True)
                towrite.seek(0)
                
                st.download_button(
                    label="🟢 Скачать итоговый анализ (Excel)",
                    data=towrite,
                    file_name="Location_Analysis_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"⚠️ Произошла непредвиденная ошибка при разборе файлов: {e}")
else:
    st.info("Пожалуйста, загрузите оба Excel-файла для глубокого факторного анализа.")
