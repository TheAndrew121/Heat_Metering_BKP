# начало файла data_entry.py
import json
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta, date


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

        # кроме 13 котельной
        boiler_options = [boiler for boiler in self.cart_weights.keys() if boiler != "Котельная №13"]
        ttk.Combobox(self, textvariable=self.boiler_var, values=boiler_options).grid(row=0, column=1, padx=5, pady=5)

        # Поле ввода даты
        ttk.Label(self, text="Дата (ДД-ММ-ГГГГ):").grid(row=1, column=0, padx=5, pady=5)
        self.date_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.date_var).grid(row=1, column=1, padx=5, pady=5)

        # Остальные поля
        labels = [
            "Расход электроэнергии", "Общий расход воды", "Расход на комплексоне",
            "Расход угля (кг)", "Оставшаяся зола", "Температура подачи воды",
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
        #ttk.Button(self, text="Настроить вес тележек", command=self.configure_cart_weights).grid(row=len(labels) + 3, column=0, columnspan=2, pady=10)
        # Кнопка для газовой котельной
        ttk.Button(self, text="Ввод данных для газовой котельной №13",
                   command=self.open_gas_boiler_window).grid(row=len(labels) + 4, column=0, columnspan=2, pady=10)

    def open_gas_boiler_window(self):
        """Открывает окно для ввода данных газовой котельной"""
        gas_window = tk.Toplevel(self)
        gas_window.title("Ввод данных для газовой котельной №13")

        # Поле ввода даты
        ttk.Label(gas_window, text="Дата (ДД-ММ-ГГГГ):").grid(row=0, column=0, padx=5, pady=5)
        date_var = tk.StringVar()
        ttk.Entry(gas_window, textvariable=date_var).grid(row=0, column=1, padx=5, pady=5)

        # Поле ввода расхода газа
        ttk.Label(gas_window, text="Расход газа (тыс. м3):").grid(row=1, column=0, padx=5, pady=5)
        gas_var = tk.StringVar()
        ttk.Entry(gas_window, textvariable=gas_var).grid(row=1, column=1, padx=5, pady=5)

        def save_gas_data():
            date = date_var.get()  # Получаем значение даты из переменной
            gas_value = gas_var.get()

            if not date or not gas_value:  # Проверяем оба поля
                messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля.")
                return

            try:
                gas_value = float(gas_value)
                datetime.strptime(date, "%d-%m-%Y")  # Проверяем формат даты
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат данных.")
                return

            # Ищем существующую запись
            existing_data = next((row for row in self.data
                                  if row["date"] == date and row["boiler"] == "Котельная №13"), None)

            if existing_data:
                existing_data["gas"] = gas_value
            else:
                new_row = {
                    "date": date,
                    "boiler": "Котельная №13",
                    "gas": gas_value,
                    "electricity": 0,
                    "total_water": 0,
                    "feed_water": 0,
                    "coal": 0,
                    "ash": 0,
                    "supply_temp": None,
                    "return_temp": None,
                    "outdoor_temp": None
                }
                self.data.append(new_row)

            self.on_save()
            messagebox.showinfo("Успех", "Данные успешно сохранены!")
            gas_window.destroy()

        ttk.Button(gas_window, text="Сохранить", command=save_gas_data).grid(row=2, column=0, columnspan=2, pady=10)


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

        # Ищем существующую запись
        existing_data = next((row for row in self.data
                              if row["date"] == date and row["boiler"] == boiler), None)

        if existing_data:
            # Проверяем каждое поле на необходимость перезаписи
            fields_to_check = {
                "electricity": "Расход электроэнергии",
                "total_water": "Общий расход воды",
                "feed_water": "Расход на комплексоне",
                "coal": "Расход угля (кг)",
                "ash": "Оставшаяся зола",
                "supply_temp": "Температура подачи воды",
                "return_temp": "Температура обратной подачи воды",
                "outdoor_temp": "Температура наружного воздуха"
            }

            for field, entry_label in fields_to_check.items():
                entry_value = self.entries[entry_label].get()
                if entry_value:  # Если поле заполнено
                    if existing_data[field] != 0:  # Если в БД уже есть значение
                        if not messagebox.askyesno(
                                "Подтверждение",
                                f"Для {date} уже есть данные по '{entry_label}'. Перезаписать?",
                                parent=self
                        ):
                            continue  # Пропускаем это поле, если диспетчер отказался

                    # Обновляем значение
                    try:
                        if field in ["coal", "ash"]:
                            existing_data[field] = float(entry_value)
                        else:
                            existing_data[field] = float(entry_value)
                    except ValueError:
                        messagebox.showerror("Ошибка", f"Неверный формат данных для '{entry_label}'")
                        return
        else:
            # Создаем новую запись
            new_row = {
                "date": date,
                "boiler": boiler,
                "electricity": float(self.entries["Расход электроэнергии"].get() or 0),
                "total_water": float(self.entries["Общий расход воды"].get() or 0),
                "feed_water": float(self.entries["Расход на комплексоне"].get() or 0),
                "coal": float(self.entries["Расход угля (кг)"].get() or 0),
                "ash": float(self.entries["Оставшаяся зола"].get() or 0),
                "supply_temp": float(self.entries["Температура подачи воды"].get() or 0),
                "return_temp": float(self.entries["Температура обратной подачи воды"].get() or 0),
                "outdoor_temp": float(self.entries["Температура наружного воздуха"].get() or 0),
            }
            self.data.append(new_row)

        # Проверка электроэнергии (если заполнено)
        if self.entries["Расход электроэнергии"].get():
            electricity_value = float(self.entries["Расход электроэнергии"].get())
            previous_data = [
                row for row in self.data
                if row["boiler"] == boiler and
                   datetime.strptime(row["date"], "%d-%m-%Y") == current_date - timedelta(days=1)
            ]
            if previous_data and electricity_value < previous_data[0]["electricity"]:
                if not messagebox.askyesno(
                        "Подтверждение",
                        f"Значение расхода электроэнергии ({electricity_value}) меньше, "
                        f"чем за предыдущий день ({previous_data[0]['electricity']}). Продолжить сохранение?",
                        parent=self
                ):
                    return

        self.on_save()
        messagebox.showinfo("Успех", "Данные успешно сохранены!")

    def configure_cart_weights(self):
        """Настройка веса тележек"""
        weight_window = tk.Toplevel(self)
        weight_window.title("Настройка веса тележек")

        entries = {}
        for i, (boiler, weights) in enumerate(self.cart_weights.items()):
            ttk.Label(weight_window, text=f"{boiler}:").grid(row=i, column=0, padx=5, pady=5)

            if boiler == "Котельная №13":
                continue
            else:
                # Для угольных котельных
                ttk.Label(weight_window, text="Уголь (кг):").grid(row=i, column=1, padx=5)
                coal_var = tk.StringVar(value=str(weights.get("coal", 0)))
                ttk.Entry(weight_window, textvariable=coal_var).grid(row=i, column=2, padx=5)

                ttk.Label(weight_window, text="Зола (кг):").grid(row=i, column=3, padx=5)
                ash_var = tk.StringVar(value=str(weights.get("ash", 0)))
                ttk.Entry(weight_window, textvariable=ash_var).grid(row=i, column=4, padx=5)

                entries[boiler] = {"coal": coal_var, "ash": ash_var}

        def save_weights():
            for boiler, vars in entries.items():
                try:
                    if boiler == "Котельная №13":
                        self.cart_weights[boiler] = {
                            "gas": float(vars["gas"].get())
                        }
                    else:
                        self.cart_weights[boiler] = {
                            "coal": float(vars["coal"].get()),
                            "ash": float(vars["ash"].get())
                        }
                except ValueError:
                    messagebox.showerror("Ошибка", f"Неверное значение веса для {boiler}.")
                    return

            # Сохраняем в файл
            try:
                with open("cart_weights.json", "w") as f:
                    json.dump(self.cart_weights, f)
                messagebox.showinfo("Успех", "Веса успешно обновлены!")
                weight_window.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить веса: {str(e)}")

        ttk.Button(weight_window, text="Сохранить", command=save_weights).grid(
            row=len(self.cart_weights), column=0, columnspan=5, pady=10)
# конец файла data_entry.py