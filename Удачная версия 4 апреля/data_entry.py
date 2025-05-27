import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta


class DataEntryFrame(ttk.LabelFrame):
    def __init__(self, parent, data, cart_weights, on_save):
        super().__init__(parent, text="Ввод данных")
        self.parent = parent
        self.data = data
        self.cart_weights = cart_weights
        self.on_save = on_save

        # Поля ввода
        self.create_widgets()

    def create_widgets(self):
        # Выбор котельной
        ttk.Label(self, text="Котельная:").grid(row=0, column=0, padx=5, pady=5)
        self.boiler_var = tk.StringVar()
        boiler_options = list(self.cart_weights.keys())
        ttk.Combobox(self, textvariable=self.boiler_var, values=boiler_options).grid(row=0, column=1, padx=5, pady=5)

        # Поле ввода даты
        ttk.Label(self, text="Дата (ДД-ММ-ГГГГ):").grid(row=1, column=0, padx=5, pady=5)
        self.date_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.date_var).grid(row=1, column=1, padx=5, pady=5)

        # Остальные поля
        labels = [
            "Расход электроэнергии", "Общий расход воды", "Расход воды на подпитку",
            "Тележки с углём", "Тележки с золой", "Температура подачи воды",
            "Температура обратной подачи воды", "Температура наружного воздуха"
        ]
        self.entries = {}
        for i, label in enumerate(labels):
            ttk.Label(self, text=label + ":").grid(row=i + 2, column=0, padx=5, pady=5)
            var = tk.StringVar()
            ttk.Entry(self, textvariable=var).grid(row=i + 2, column=1, padx=5, pady=5)
            self.entries[label] = var

        # Кнопка сохранения
        ttk.Button(self, text="Сохранить данные", command=self.save_data).grid(row=len(labels) + 2, column=0, columnspan=2, pady=10)

        # Кнопка настройки веса тележек
        ttk.Button(self, text="Настроить вес тележек", command=self.configure_cart_weights).grid(row=len(labels) + 3, column=0, columnspan=2, pady=10)

    def save_data(self):
        date = self.date_var.get()
        boiler = self.boiler_var.get()
        if not date or not boiler:
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все обязательные поля.")
            return

        try:
            current_date = datetime.strptime(date, "%d-%m-%Y")
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД-ММ-ГГГГ.")
            return

        # Проверка на наличие данных за предыдущий день
        previous_data = [
            row for row in self.data
            if
            row["boiler"] == boiler and datetime.strptime(row["date"], "%d-%m-%Y") == current_date - timedelta(days=1)
        ]

        electricity_value = float(self.entries["Расход электроэнергии"].get() or 0)

        if previous_data:
            previous_electricity = previous_data[0]["electricity"]
            if electricity_value < previous_electricity:
                messagebox.showerror(
                    "Ошибка",
                    f"Значение расхода электроэнергии ({electricity_value}) меньше, чем за предыдущий день ({previous_electricity})."
                )
                return

        # Создание новой строки данных
        new_row = {
            "date": date,
            "boiler": boiler,
            "electricity": electricity_value,
            "total_water": float(self.entries["Общий расход воды"].get() or 0),
            "feed_water": float(self.entries["Расход воды на подпитку"].get() or 0),
            "coal_carts": int(self.entries["Тележки с углём"].get() or 0),
            "ash_carts": int(self.entries["Тележки с золой"].get() or 0),
            "supply_temp": float(self.entries["Температура подачи воды"].get() or 0),
            "return_temp": float(self.entries["Температура обратной подачи воды"].get() or 0),
            "outdoor_temp": float(self.entries["Температура наружного воздуха"].get() or 0),
        }

        # Добавляем новую строку в данные
        self.data.append(new_row)
        self.on_save()

    def configure_cart_weights(self):
        """Настройка веса тележек"""
        weight_window = tk.Toplevel(self)
        weight_window.title("Настройка веса тележек")

        for i, (boiler, weight) in enumerate(self.cart_weights.items()):
            ttk.Label(weight_window, text=f"{boiler}:").grid(row=i, column=0, padx=5, pady=5)
            var = tk.StringVar(value=str(weight))
            ttk.Entry(weight_window, textvariable=var).grid(row=i, column=1, padx=5, pady=5)
            self.cart_weights[boiler] = var

        def save_weights():
            for boiler, var in self.cart_weights.items():
                try:
                    self.cart_weights[boiler] = float(var.get())
                except ValueError:
                    messagebox.showerror("Ошибка", f"Неверное значение веса для {boiler}.")
                    return
            messagebox.showinfo("Успех", "Веса тележек успешно обновлены!")
            weight_window.destroy()

        ttk.Button(weight_window, text="Сохранить", command=save_weights).grid(row=len(self.cart_weights), column=0, columnspan=2, pady=10)