from tkinter import filedialog, messagebox
import tkinter as tk
import pandas as pd
import re
from openpyxl import load_workbook
from docx import Document
import os
#import difflib



# Словарь синонимов
SYNONYMS = {
    "Мощность, Вт": ["мощность", "энергопотребление", "Вт", "W"],
    "Св. поток, Лм": ["световой поток", "Лм", "Lm"],
    "IP": ["степень защиты", "IP"],
    "Длина, мм": ["длина", "L"],
    "Ширина, мм": ["ширина", "B"],
    "Высота, мм": ["высота", "H", "h"],
    "Гарантия": ["гарантийный срок", "срок гарантии"],
    "Прочее": []
}

FORM_2_COLUMNS = [
    "Номенклатура", "Мощность, Вт", "Св. поток, Лм", "IP", "Габариты, мм (L,b,h)",
    "Длина, мм", "Ширина, мм", "Высота, мм", "Рассеиватель", "Цвет. температура, К",
    "Вес, кг", "Напряжение, В", "Температура эксплуатации", "Срок службы/работы светильника",
    "Тип КСС", "Род тока", "Гарантия", "Индекс цветопередачи, CRI, Ra",
    "Цвет корпуса", "Коэффициент пульсаций", "Коэффициент мощности, Pf",
    "Класс взрывозащиты, Ex", "Класс пожароопасности",
    "Класс защиты от поражения электрическим током", "Материал корпуса", "Тип", "Прочее"
]


import re

def extract_params(cell_value):
    """
    Извлечение параметров и значений из строки полного описания.
    """
    result = {}
    try:
        # Расширенные шаблоны
        patterns = [
            (r"(потребляемая мощность|мощность)\s*[:\-–]?\s*(\d+\s*Вт)", "Мощность, Вт"),
            (r"(световой поток)\s*[:\-–]?\s*(\d+\s*лм)", "Св. поток, Лм"),
            (r"(цветовая температура|Цвет. температура, К)\s*[:\-–]?\s*(\d+\s*[КK])", "Цвет. температура, К"),
            (r"(степень защиты|IP)\s*[:\-–]?\s*(IP\d+)", "IP"),
            (r"(размеры|габариты)\s*[:\-–]?\s*(\d+\s*[хx]\s*\d+\s*[хx]\s*\d+)", "Габариты, мм (L,b,h)"),
            (r"(напряжение питания|Напряжение, В)\s*[:\-–]?\s*([\d\-–]+\s*[Вv])", "Напряжение, В"),
            (r"(материалы|Материал корпуса)\s*[:\-–]?\s*(.+?)(?:,|$)", "Материал корпуса"),
            (r"(серия)\s*[:\-–]?\s*(РИСТ-[\w\-]+)", "Номенклатура"),
            (r"(климатического исполнения|исполнение)\s*[:\-–]?\s*(УХЛ\d+)", "Прочее"),
        ]

        for pattern, column_name in patterns:
            match = re.search(pattern, str(cell_value), re.IGNORECASE)
            if match:
                result[column_name] = match.group(2).strip()

    except Exception as e:
        result['error'] = f"Ошибка обработки строки: {e}"
    return result


def map_to_columns(params):
    """
    Сопоставление параметров с колонками таблицы "Форма 2".
    """
    mapped_data = {key: None for key in FORM_2_COLUMNS}
    for param, value in params.items():
        if param in FORM_2_COLUMNS:
            mapped_data[param] = value
        else:
            # Если параметр не соответствует колонкам, добавляем в "Прочее"
            if value:
                mapped_data["Прочее"] = (mapped_data.get("Прочее", "") or "") + f"{param}: {value}; "
    return mapped_data





def adjust_column_width(output_path):
    """
    Устанавливает ширину колонок в Excel-файле на основе их содержимого.
    """
    try:
        workbook = load_workbook(output_path)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Получаем букву колонки
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2  # Добавляем небольшой запас
            sheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(output_path)
        workbook.close()
    except Exception as e:
        print(f"Ошибка при изменении ширины колонок: {e}")


def process_file(input_path, output_path):
    """
    Обработка Excel или Word файлов с длинными описаниями.
    """
    try:
        file_extension = os.path.splitext(input_path)[1].lower()

        if file_extension == ".xlsx":
            data = pd.read_excel(input_path)
        elif file_extension == ".docx":
            word_data = read_word_file(input_path)
            data = pd.DataFrame({"Наименование позиции": word_data})
        else:
            raise ValueError("Поддерживаются только файлы Excel (.xlsx) и Word (.docx).")

        if file_extension == ".xlsx":
            data.columns = data.columns.str.strip()

        if "Наименование позиции" not in data.columns:
            raise ValueError("Колонка 'Наименование позиции' не найдена в файле!")

        result_df = pd.DataFrame(columns=FORM_2_COLUMNS)

        for _, row in data.iterrows():
            full_description = row.get("Наименование позиции")
            if pd.notna(full_description):
                extracted = extract_params(full_description)
                print("Извлеченные параметры:", extracted)  # Вывод для проверки
                mapped = map_to_columns(extracted)
                result_df = pd.concat([result_df, pd.DataFrame([mapped])], ignore_index=True)

        result_df.to_excel(output_path, index=False)
        adjust_column_width(output_path)
        messagebox.showinfo("Готово", f"Файл обработан и сохранен в: {output_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def select_file():
    """
    Выбор файла для обработки.
    """
    input_path = filedialog.askopenfilename(
        title="Выберите файл",
        filetypes=[("Supported Files", "*.xlsx *.docx"), ("Excel Files", "*.xlsx"), ("Word Files", "*.docx")]
    )
    if not input_path:
        return

    output_path = filedialog.asksaveasfilename(
        title="Сохранить как",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_path:
        return

    process_file(input_path, output_path)


def read_word_file(file_path):
    """
    Извлекает текст из Word-файла (.docx) постранично.
    """
    document = Document(file_path)
    rows = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:  # Пропускаем пустые строки
            rows.append(text)
    return rows


# Интерфейс
root = tk.Tk()
root.title("Обработка файлов")
root.geometry("400x200")

label = tk.Label(root, text="Выберите Excel или Word-файл для обработки:")
label.pack(pady=10)

btn_select = tk.Button(root, text="Выбрать файл", command=select_file)
btn_select.pack(pady=20)

root.mainloop()

