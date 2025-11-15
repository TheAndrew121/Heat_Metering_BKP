# начало файла Boiler.py
import json
import tkinter as tk
from tkinter import ttk, messagebox
from data_entry import DataEntryFrame
from report_generator import ReportGeneratorFrame
from database import DatabaseManager
from anomaly_detection import detect_anomalies

# Глобальные переменные
DB_FILE = "boiler_data.db"

def load_cart_weights():
    try:
        with open("cart_weights.json", "r", encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        default_weights = {
            "Котельная №6": {"coal": 83, "ash": 58},
            "Котельная №9": {"coal": 77, "ash": 50},
            "Котельная №11": {"coal": 92, "ash": 50},
            "Котельная №12": {"coal": 79, "ash": 34},
            "Котельная №15": {"coal": 88, "ash": 101},
            "Котельная №17": {"coal": 86, "ash": 50},
            "Котельная №18": {"coal": 84, "ash": 54},
            "Котельная №13": {"gas": 1},
        }
        with open("cart_weights.json", "w", encoding='utf-8') as f:
            json.dump(default_weights, f, ensure_ascii=False)
        return default_weights


class BoilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт данных котельных")
        self.db_manager = DatabaseManager(DB_FILE)
        self.cart_weights = load_cart_weights()

        # Загрузка данных
        self.data = self.db_manager.load_data()
        self.rek_limits = self.load_rek_limits()  # Загрузка данных РЭК

        # Создание фреймов
        self.data_entry_frame = DataEntryFrame(self.root, self.data, self.cart_weights, self.save_data)

        self.report_generator_frame = ReportGeneratorFrame(
            self.root,  # Родительский виджет
            self.data,
            self.cart_weights,
            lambda data: self.detect_anomalies(data),
            self  # Ссылка на BoilerApp для доступа к rek_limits
        )


        # Размещение фреймов
        self.data_entry_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.report_generator_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        """РЭК – региональная энергетическая комиссия. 
        Данные по РЭК определяют, 
        сколько котельной разрешено потратить того или иного ресурса"""
        # Ввод данных РЭК
        ttk.Button(self.root, text="Настроить нормативы РЭК", command=self.configure_rek_limits).grid(row=1, column=0, columnspan=2,
                                                                                                      pady=10)
        # Настройка геометрии окна
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

    @staticmethod
    def load_rek_limits():
        """Загрузка данных РЭК"""
        with open("rek_limits.json", "r", encoding='utf-8') as f:
            return json.load(f)

    def save_rek_limits(self):
        """Сохранение данных РЭК"""
        with open("rek_limits.json", "w") as f:
            json.dump(self.rek_limits, f)


    def configure_rek_limits(self):
        """Настройка нормативов РЭК"""
        rek_window = tk.Toplevel(self.root)
        rek_window.title("Настройка нормативов РЭК")

        # Создаем Notebook для вкладок
        notebook = ttk.Notebook(rek_window)
        notebook.pack(padx=10, pady=10, fill='both', expand=True)

        # Создаем фреймы для каждой котельной
        entries = {}
        for boiler_id, boiler_data in self.rek_limits["boilers"].items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=f"Котельная №{boiler_id}")

            # Периоды: первое и второе полугодие
            periods = ["first_half", "second_half"]
            period_labels = ["Первое полугодие", "Второе полугодие"]

            for period_idx, period in enumerate(periods):
                ttk.Label(frame, text=period_labels[period_idx], font=('Arial', 10, 'bold')).grid(
                    row=period_idx * 5, column=0, columnspan=2, pady=(10, 5))

                labels = ["Уголь (т)", "Электроэнергия (кВт·ч)", "Тепло (Гкал)", "Вода (м³)"]
                fields = ["coal", "electricity", "gcal", "water"]

                for i, (label, field) in enumerate(zip(labels, fields)):
                    ttk.Label(frame, text=label + ":").grid(
                        row=period_idx * 5 + i + 1, column=0, padx=5, pady=2, sticky='e')

                    var = tk.StringVar(value=str(boiler_data[period][field]))
                    ttk.Entry(frame, textvariable=var, width=15).grid(
                        row=period_idx * 5 + i + 1, column=1, padx=5, pady=2)

                    if boiler_id not in entries:
                        entries[boiler_id] = {}
                    if period not in entries[boiler_id]:
                        entries[boiler_id][period] = {}

                    entries[boiler_id][period][field] = var

        # Газовая котельная и коэффициенты по месяцам
        gas_frame = ttk.Frame(notebook)
        notebook.add(gas_frame, text="Газовая котельная")

        # Газ для первого и второго полугодия
        ttk.Label(gas_frame, text="Газ (первое полугодие, тыс. м³):").grid(row=0, column=0, padx=5, pady=5)
        gas_first_var = tk.StringVar(value=str(self.rek_limits.get("gas_first_half", 0)))
        ttk.Entry(gas_frame, textvariable=gas_first_var).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(gas_frame, text="Газ (второе полугодие, тыс. м³):").grid(row=1, column=0, padx=5, pady=5)
        gas_second_var = tk.StringVar(value=str(self.rek_limits.get("gas_second_half", 0)))
        ttk.Entry(gas_frame, textvariable=gas_second_var).grid(row=1, column=1, padx=5, pady=5)

        # Коэффициенты по месяцам
        coeff_frame = ttk.Frame(notebook)
        notebook.add(coeff_frame, text="Коэффициенты по месяцам")

        month_names = {
            "1": "Январь", "2": "Февраль", "3": "Март", "4": "Апрель",
            "5": "Май", "9": "Сентябрь", "10": "Октябрь", "11": "Ноябрь", "12": "Декабрь"
        }

        month_entries = {}
        for i, (month_num, month_name) in enumerate(month_names.items()):
            row = i // 2
            col = (i % 2) * 2

            ttk.Label(coeff_frame, text=month_name + ":").grid(row=row, column=col, padx=5, pady=2, sticky='e')
            var = tk.StringVar(value=str(self.rek_limits["month_coefficients"].get(month_num, 0)))
            ttk.Entry(coeff_frame, textvariable=var, width=8).grid(row=row, column=col + 1, padx=5, pady=2)
            month_entries[month_num] = var

        def save_limits():
            # Сохраняем данные по котельным
            for boiler_id, period_data in entries.items():
                for period, field_data in period_data.items():
                    for field, var in field_data.items():
                        try:
                            self.rek_limits["boilers"][boiler_id][period][field] = float(var.get())
                        except ValueError:
                            messagebox.showerror("Ошибка",
                                                 f"Неверное значение для котельной {boiler_id}, {period}, {field}.")
                            return

            # Сохраняем данные по газу
            try:
                self.rek_limits["gas_first_half"] = float(gas_first_var.get())
                self.rek_limits["gas_second_half"] = float(gas_second_var.get())
            except ValueError:
                messagebox.showerror("Ошибка", "Неверное значение для газа.")
                return

            # Сохраняем коэффициенты по месяцам
            for month_num, var in month_entries.items():
                try:
                    self.rek_limits["month_coefficients"][month_num] = float(var.get())
                except ValueError:
                    messagebox.showerror("Ошибка", f"Неверное значение коэффициента для месяца {month_num}.")
                    return

            self.save_rek_limits()
            messagebox.showinfo("Успех", "Нормативы РЭК успешно обновлены!")
            rek_window.destroy()

        ttk.Button(rek_window, text="Сохранить все настройки", command=save_limits).pack(pady=10)

    def detect_anomalies(self, filtered_data):
        """Выявление аномалий"""
        return detect_anomalies(filtered_data, self.cart_weights, self.rek_limits)

    def save_data(self):
        """Сохранение данных"""
        self.db_manager.save_data(self.data)


if __name__ == "__main__":
    root = tk.Tk()
    app = BoilerApp(root)
    root.mainloop()
# конец файла Boiler.py