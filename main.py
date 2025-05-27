import json
import tkinter as tk
from tkinter import ttk, messagebox
from data_entry import DataEntryFrame
from report_generator import ReportGeneratorFrame
from database import DatabaseManager
from anomaly_detection import detect_anomalies

# Глобальные переменные
DB_FILE = "boiler_data.db"
DEFAULT_CART_WEIGHTS = {
    "Котельная №6": 111,
    "Котельная №9": 111,
    "Котельная №11": 111,
    "Котельная №12": 111,
    "Котельная №15": 111,
    "Котельная №17": 111,
    "Котельная №18": 111,
}

class BoilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт данных котельных")
        self.db_manager = DatabaseManager(DB_FILE)
        self.cart_weights = DEFAULT_CART_WEIGHTS.copy()

        # Загрузка данных
        self.data = self.db_manager.load_data()
        self.rek_limits = self.load_rek_limits()  # Загрузка данных РЭК

        # Создание фреймов
        self.data_entry_frame = DataEntryFrame(self.root, self.data, self.cart_weights, self.save_data)
        # В классе BoilerApp изменить строку создания ReportGeneratorFrame:
        self.report_generator_frame = ReportGeneratorFrame(
            self.root,
            self.data,
            self.cart_weights,
            lambda data: self.detect_anomalies(data))

        # Ввод данных РЭК
        ttk.Button(self.root, text="Настроить нормативы РЭК", command=self.configure_rek_limits).grid(row=1, column=0, columnspan=2, pady=10)

        # Размещение фреймов
        self.data_entry_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.report_generator_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # Настройка геометрии окна
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

    @staticmethod
    def load_rek_limits():
        """Загрузка данных РЭК"""
        try:
            with open("rek_limits.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return {"electricity": 1000, "total_water": 500, "feed_water": 200}

    def save_rek_limits(self):
        """Сохранение данных РЭК"""
        with open("rek_limits.json", "w") as f:
            json.dump(self.rek_limits, f)

    def configure_rek_limits(self):
        """Настройка нормативов РЭК"""
        rek_window = tk.Toplevel(self.root)
        rek_window.title("Настройка нормативов РЭК")

        labels = ["electricity", "total_water", "feed_water"]
        entries = {}
        for i, label in enumerate(labels):
            ttk.Label(rek_window, text=label.capitalize() + ":").grid(row=i, column=0, padx=5, pady=5)
            var = tk.StringVar(value=str(self.rek_limits[label]))
            ttk.Entry(rek_window, textvariable=var).grid(row=i, column=1, padx=5, pady=5)
            entries[label] = var

        def save_limits():
            for label, var in entries.items():
                try:
                    self.rek_limits[label] = float(var.get())
                except ValueError:
                    messagebox.showerror("Ошибка", f"Неверное значение для {label}.")
                    return
            self.save_rek_limits()
            messagebox.showinfo("Успех", "Нормативы РЭК успешно обновлены!")
            rek_window.destroy()

        ttk.Button(rek_window, text="Сохранить", command=save_limits).grid(row=len(labels), column=0, columnspan=2, pady=10)

    def detect_anomalies(self, filtered_data):
        """Выявление аномалий"""
        return detect_anomalies(filtered_data, self.cart_weights, self.rek_limits)

    def save_data(self):
        """Сохранение данных"""
        self.db_manager.save_data(self.data)
        messagebox.showinfo("Успех", "Данные успешно сохранены!")

if __name__ == "__main__":
    root = tk.Tk()
    app = BoilerApp(root)
    root.mainloop()