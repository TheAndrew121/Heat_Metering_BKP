# начало файла anomaly_detection.py
from datetime import datetime, timedelta


def detect_anomalies(data, cart_weights, rek_limits):
    anomalies = []

    for idx, row in enumerate(data):
        issues = []

        boiler = row.get("boiler", "")
        if boiler == "Котельная №13":
            continue

        # Проверка резкого изменения для электричества: скачок больше 14.2 кВт/ч за сутки - аномалия
        if idx > 0 and abs(row["electricity"] - data[idx - 1]["electricity"]) > 14.2:
            issues.append("electricity")

        # Проверка резкого изменения для общей воды: скачок больше 0.1 кубометра за сутки - аномалия
        if idx > 0 and abs(row["total_water"] - data[idx - 1]["total_water"]) > 0.1:
            issues.append("total_water")

        # Проверка резкого изменения для комплексона: скачок больше 0.1 кубометра за сутки - аномалия
        if idx > 0 and abs(row["feed_water"] - data[idx - 1]["feed_water"]) > 0.1:
            issues.append("feed_water")

        # ===================================================================================================================

        # Проверка неизменности данных для электричества: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["electricity"] == row["electricity"] for i in range(idx - 2, idx)):
            issues.append("electricity")

        # Проверка неизменности данных для воды: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["total_water"] == row["total_water"] for i in range(idx - 2, idx)):
            issues.append("total_water")

        # Проверка неизменности данных для комплексона: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["feed_water"] == row["feed_water"] for i in range(idx - 2, idx)):
            issues.append("feed_water")

        # Проверка температур по графику
        temp_graph = load_temperature_graph()
        expected_temps = temp_graph.get(row["outdoor_temp"], {})

        if expected_temps:
            expected_supply_temp = expected_temps.get("supply_temp")
            expected_return_temp = expected_temps.get("return_temp")

            # Проверка температуры воды в подающем трубопроводе
            if expected_supply_temp is not None and row["supply_temp"] != expected_supply_temp:
                issues.append("supply_temp")

            # Проверка температуры воды в обратном трубопроводе
            if expected_return_temp is not None and row["return_temp"] != expected_return_temp:
                issues.append("return_temp")

        if issues:
            anomalies.append({idx: issues})

    return anomalies

# Загрузка данных из файла Температурный график.xlsx
def load_temperature_graph():
    import pandas as pd

    try:
        graph = pd.read_excel("Температурный график.xlsx")
        temp_graph = {}
        for _, row in graph.iterrows():
            temp_graph[row["Температура наружного воздуха"]] = {
                "supply_temp": row["Температура воды в подающем трубопроводе"],
                "return_temp": row["Температура воды в обратном трубопроводе"],
            }
        return temp_graph
    except Exception as e:
        print(f"Ошибка загрузки температурного графика: {e}")
        return {}
# конец файла anomaly_detection.py