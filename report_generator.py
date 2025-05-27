import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re


def sanitize_filename(filename):
    """Удаляет недопустимые символы из имени файла."""
    return re.sub(r'[<>:"/\\|?*]', '', filename)


class ReportGeneratorFrame(ttk.LabelFrame):
    def __init__(self, parent, data, cart_weights, detect_anomalies):
        super().__init__(parent, text="Формирование отчётов")
        self.parent = parent
        self.data = data
        self.cart_weights = cart_weights
        self.detect_anomalies = detect_anomalies

        self.create_widgets()

    def create_widgets(self):
        # Выбор котельной
        ttk.Label(self, text="Котельная:").grid(row=0, column=0, padx=5, pady=5)
        self.boiler_var = tk.StringVar()
        boiler_options = list(self.cart_weights.keys())
        ttk.Combobox(self, textvariable=self.boiler_var, values=boiler_options).grid(row=0, column=1, padx=5, pady=5)

        # Выбор периода
        ttk.Label(self, text="Начальная дата:").grid(row=1, column=0, padx=5, pady=5)
        self.start_date_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.start_date_var).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self, text="Конечная дата:").grid(row=2, column=0, padx=5, pady=5)
        self.end_date_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.end_date_var).grid(row=2, column=1, padx=5, pady=5)

        # Выбор типа отчёта
        ttk.Label(self, text="Тип отчёта:").grid(row=3, column=0, padx=5, pady=5)
        self.report_type_var = tk.StringVar()

        report_types = [
            "Полный отчёт",
            "Температуры",
            "Расход тележек",
            "Расход воды",
            "Расход комплексона",
            "Расход электроэнергии"
        ]
        ttk.Combobox(self, textvariable=self.report_type_var, values=report_types).grid(row=3, column=1, padx=5, pady=5)

        # Кнопка формирования отчёта
        ttk.Button(self, text="Сформировать отчёт", command=self.generate_report).grid(row=4, column=0, columnspan=2,
                                                                                       pady=10)

    def generate_report(self):
        boiler = self.boiler_var.get()
        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()
        report_type = self.report_type_var.get()

        if not boiler or not start_date or not end_date or not report_type:
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля.")
            return

        try:
            start_date_dt = pd.to_datetime(start_date, format="%d-%m-%Y")
            end_date_dt = pd.to_datetime(end_date, format="%d-%m-%Y")
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД-ММ-ГГГГ.")
            return

        # Фильтрация данных по котельной и дате
        filtered_data = [
            row for row in self.data
            if (row["boiler"] == boiler and
                start_date_dt <= pd.to_datetime(row["date"], format="%d-%m-%Y") <= end_date_dt)
        ]

        if not filtered_data:
            messagebox.showinfo("Информация", f"Нет данных для котельной {boiler} за указанный период.")
            return

        # Обнаружение аномалий
        anomalies = self.detect_anomalies(filtered_data)

        # Создание Excel-файла
        wb = Workbook()
        ws = wb.active
        ws.title = "Отчёт"

        # Определение заголовков в зависимости от типа отчёта
        if report_type == "Полный отчёт":
            headers = ["Дата", "Котельная", "Расход электроэнергии", "Общий расход воды",
                       "Расход комплексона", "Тележки с углём", "Тележки с золой",
                       "Температура подачи воды", "Температура обратной подачи воды",
                       "Температура наружного воздуха"]
        elif report_type == "Температуры":
            headers = ["Дата", "Котельная", "Температура подачи воды",
                       "Температура обратной подачи воды", "Температура наружного воздуха"]
        elif report_type == "Расход тележек":
            headers = ["Дата", "Котельная", "Тележки с углём", "Тележки с золой"]
        elif report_type == "Расход воды":
            headers = ["Дата", "Котельная", "Общий расход воды"]
        elif report_type == "Расход комплексона":
            headers = ["Дата", "Котельная", "комплексона"]
        elif report_type == "Расход электроэнергии":
            headers = ["Дата", "Котельная", "Расход электроэнергии"]

        ws.append(headers)

        # Словарь для сопоставления английских ключей с русскими названиями столбцов
        key_to_header = {
            "date": "Дата",
            "boiler": "Котельная",
            "electricity": "Расход электроэнергии",
            "total_water": "Общий расход воды",
            "feed_water": "Расход комплексона",
            "coal_carts": "Тележки с углём",
            "ash_carts": "Тележки с золой",
            "supply_temp": "Температура подачи воды",
            "return_temp": "Температура обратной подачи воды",
            "outdoor_temp": "Температура наружного воздуха"
        }

        # Заполнение данных
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        for row_idx, row in enumerate(filtered_data, start=2):
            if report_type == "Полный отчёт":
                data_row = [
                    row["date"], row["boiler"], row["electricity"], row["total_water"],
                    row["feed_water"], row["coal_carts"], row["ash_carts"],
                    row["supply_temp"], row["return_temp"], row["outdoor_temp"]
                ]
            elif report_type == "Температуры":
                data_row = [
                    row["date"], row["boiler"], row["supply_temp"],
                    row["return_temp"], row["outdoor_temp"]
                ]
            elif report_type == "Расход тележек":
                data_row = [
                    row["date"], row["boiler"], row["coal_carts"], row["ash_carts"]
                ]
            elif report_type == "Расход воды":
                data_row = [
                    row["date"], row["boiler"], row["total_water"]
                ]
            elif report_type == "Расход комплексона":
                data_row = [
                    row["date"], row["boiler"], row["feed_water"]
                ]
            elif report_type == "Расход электроэнергии":
                data_row = [
                    row["date"], row["boiler"], row["electricity"]
                ]

            ws.append(data_row)

            # Подсветка аномалий
            for anomaly in anomalies:
                for anomaly_row_idx, anomaly_keys in anomaly.items():
                    if row_idx == anomaly_row_idx + 2:
                        for key in anomaly_keys:
                            header = key_to_header.get(key)
                            if header and header in headers:
                                col_idx = headers.index(header) + 1
                                ws.cell(row=row_idx, column=col_idx).fill = red_fill

        # Сохранение файла
        sanitized_boiler_name = sanitize_filename(boiler)
        sanitized_report_type = sanitize_filename(report_type)
        file_name = f"{sanitized_boiler_name} {sanitized_report_type} {start_date} - {end_date}.xlsx"
        reports_dir = "reports"
        os.makedirs(reports_dir, exist_ok=True)
        file_path = os.path.join(reports_dir, file_name)
        wb.save(file_path)
        messagebox.showinfo("Успешно!", f"Отчёт сохранён в файл: {file_path}")