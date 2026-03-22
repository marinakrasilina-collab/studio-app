import pandas as pd
from io import BytesIO
import datetime

# 1. Создадим тестовые данные (имитация того, что вы ввели бы в приложении)
data = [
    {'date': '2024-10-25', 'time': '10:00', 'staff': 'Катя', 'service': 'Запись на студию', 'packet': 'LITE', 'client_name': 'Иванова Анна', 'phone': '+79001112233', 'duration': 1.0, 'status': 'Предоплата+'},
    {'date': '2024-10-25', 'time': '12:00', 'staff': 'Женя', 'service': 'Урок по вокалу (Абонемент)', 'packet': 'STANDARD', 'client_name': 'Петров Олег', 'phone': '+79004445566', 'duration': 1.5, 'status': 'Оплачено'},
    {'date': '2024-10-25', 'time': '15:00', 'staff': 'Юля', 'service': 'Пробный урок', 'packet': 'РАЗОВОЕ', 'client_name': 'Сидорова Мария', 'phone': '+79007778899', 'duration': 0.5, 'status': 'Ожидает'},
    # Добавьте сюда свои строки вручную, если нужно протестировать
]

df = pd.DataFrame(data)

# 2. Функция создания красивого Excel
def create_excel(df):
    output = BytesIO()
    
    # Сортировка
    df_sorted = df.sort_values(by=['date', 'time'])
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sorted.to_excel(writer, index=False, sheet_name='Расписание')
        
        workbook = writer.book
        worksheet = writer.sheets['Расписание']
        
        # Стили
        from openpyxl.styles import Font, PatternFill, Alignment
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        # Красим заголовки
        for cell in worksheet["1:1"]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            
        # Красим строки по сотрудникам
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, values_only=False), start=2):
            staff_name = str(row[2].value) # Столбец C - сотрудник
            fill = None
            font_color = "000000"
            
            if 'Катя' in staff_name:
                fill = PatternFill(start_color="3b82f6", end_color="3b82f6", fill_type="solid")
                font_color = "FFFFFF"
            elif 'Женя' in staff_name:
                fill = PatternFill(start_color="22c55e", end_color="22c55e", fill_type="solid")
                font_color = "FFFFFF"
            elif 'Юля' in staff_name:
                fill = PatternFill(start_color="f97316", end_color="f97316", fill_type="solid")
                font_color = "FFFFFF"
                
            if fill:
                for cell in row:
                    cell.fill = fill
                    cell.font = Font(color=font_color, bold=True)
                    cell.alignment = Alignment(horizontal='center')
        
        # Автоширина
        for col in worksheet.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            worksheet.column_dimensions[col[0].column_letter].width = min(max_len + 2, 25)

    return output.getvalue()

# 3. Генерация и сохранение
excel_data = create_excel(df)

# В большинстве онлайн-редакторов файл сохраняется во временную папку
file_name = "Studio_Report_Test.xlsx"
with open(file_name, "wb") as f:
    f.write(excel_data)

print(f"✅ Файл '{file_name}' успешно создан!")
print("👉 Ищите его во вкладке 'Files' или скачайте через интерфейс сервиса.")
