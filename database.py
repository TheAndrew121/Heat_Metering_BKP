# начало файла database.py

import sqlite3
class DatabaseManager:
    def __init__(self, db_file):
        self.conn = sqlite3.connect(db_file)
        self.conn.execute("PRAGMA encoding = 'UTF-8'")
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        """таблица для хранения данных"""
        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS boiler_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            boiler TEXT,
            electricity REAL,
            total_water REAL,
            feed_water REAL,
            coal REAL,
            ash REAL,
            supply_temp REAL,
            return_temp REAL,
            outdoor_temp REAL,
            gas REAL
        )
        """)
        self.conn.commit()

    def load_data(self):
        """Загрузка данных из базы данных"""
        self.cursor.execute("SELECT * FROM boiler_data")
        rows = self.cursor.fetchall()
        columns = [desc[0] for desc in self.cursor.description]
        data = []
        for row in rows:
            data.append({columns[i]: row[i] for i in range(len(columns))})
        return data

    def save_data(self, data):
        for row in data:
            # Проверяем, что date и boiler - строки
            if not isinstance(row.get("date", ""), str) or not isinstance(row.get("boiler", ""), str):
                continue

            # Проверка на существующую запись
            self.cursor.execute("SELECT * FROM boiler_data WHERE date = ? AND boiler = ?",
                                (str(row["date"]), str(row["boiler"])))  # Явное преобразование в строку
            exists = self.cursor.fetchone()

            if exists:
                # Обновляем существующую запись
                self.cursor.execute("""
                UPDATE boiler_data SET
                    electricity = ?,
                    total_water = ?,
                    feed_water = ?,
                    coal = ?,
                    ash = ?,
                    supply_temp = ?,
                    return_temp = ?,
                    outdoor_temp = ?,
                    gas = ?
                WHERE date = ? AND boiler = ?
                """, (
                    row.get("electricity", 0),
                    row.get("total_water", 0),
                    row.get("feed_water", 0),
                    row.get("coal", 0),
                    row.get("ash", 0),
                    row.get("supply_temp"),
                    row.get("return_temp"),
                    row.get("outdoor_temp"),
                    row.get("gas"),
                    str(row["date"]),
                    str(row["boiler"])
                ))
            else:
                # Вставляем новую запись
                self.cursor.execute("""
                INSERT INTO boiler_data 
                (date, boiler, electricity, total_water, feed_water, coal, ash, 
                 supply_temp, return_temp, outdoor_temp, gas)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    str(row["date"]),
                    str(row["boiler"]),
                    row.get("electricity", 0),
                    row.get("total_water", 0),
                    row.get("feed_water", 0),
                    row.get("coal", 0),
                    row.get("ash", 0),
                    row.get("supply_temp"),
                    row.get("return_temp"),
                    row.get("outdoor_temp"),
                    row.get("gas")
                ))
        self.conn.commit()
# конец файла database.py