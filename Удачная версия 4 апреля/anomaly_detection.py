def detect_anomalies(data, cart_weights, rek_limits):
    anomalies = []
    for idx, row in enumerate(data):
        issues = []

        # Проверка превышения лимита РЭК
        if row["electricity"] > rek_limits["electricity"]:
            issues.append("electricity")
        if row["total_water"] > rek_limits["total_water"]:
            issues.append("total_water")
        if row["feed_water"] > rek_limits["feed_water"]:
            issues.append("feed_water")

        # Проверка резкого изменения для электричества: скачок больше 14.2 кВт/ч за сутки - аномалия
        if idx > 0 and abs(row["electricity"] - data[idx - 1]["electricity"]) > 14.2:
            issues.append("electricity")


        # Проверка резкого изменения для общей воды: скачок больше 0.1 кубометра за сутки - аномалия
        if idx > 0 and abs(row["total_water"] - data[idx - 1]["total_water"]) > 0.1:
            issues.append("total_water")


        # Проверка резкого изменения для комплексона: скачок больше 0.1 кубометра за сутки - аномалия
        if idx > 0 and abs(row["feed_water"] - data[idx - 1]["feed_water"]) > 0.1:
            issues.append("feed_water")

    #===================================================================================================================

        # Проверка неизменности данных для электричества: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["electricity"] == row["electricity"] for i in range(idx - 2, idx)):
            issues.append("electricity")


        # Проверка неизменности данных для воды: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["electricity"] == row["electricity"] for i in range(idx - 2, idx)):
            issues.append("electricity")


        # Проверка неизменности данных для комплексона: больше 3 ячеек неизменных данных подряд - аномалия
        if idx > 2 and all(data[i]["electricity"] == row["electricity"] for i in range(idx - 2, idx)):
            issues.append("electricity")


        # Проверка температур по графику
        temp_graph = load_temperature_graph()
        expected_supply_temp = temp_graph.get(row["outdoor_temp"], {}).get("supply_temp")
        expected_return_temp = temp_graph.get(row["outdoor_temp"], {}).get("return_temp")
        if expected_supply_temp and row["supply_temp"] < expected_supply_temp:
            issues.append("supply_temp")
        if expected_return_temp and row["return_temp"] != expected_return_temp:
            issues.append("return_temp")

        if issues:
            anomalies.append({idx: issues})

    return anomalies

# Загрузка данных из файла "Температурный график.xlsx"
def load_temperature_graph():

    import pandas as pd

    graph = pd.read_excel("Температурный график.xlsx")
    temp_graph = {}
    for _, row in graph.iterrows():
        temp_graph[row["Температура наружного воздуха"]] = {
            "supply_temp": row["Температура воды в подающем трубопроводе"],
            "return_temp": row["Температура воды в обратном трубопроводе"],
        }
    return temp_graph