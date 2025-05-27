import sqlite3

class DatabaseManager:
    def __init__(self, db_file):
        self.conn = sqlite3.connect(db_file)
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
            coal_carts INTEGER,
            ash_carts INTEGER,
            supply_temp REAL,
            return_temp REAL,
            outdoor_temp REAL
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
        """Сохранение данных в базу данных"""
        self.cursor.execute("DELETE FROM boiler_data")  # Очистка старых данных
        for row in data:
            self.cursor.execute("""
            INSERT INTO boiler_data (date, boiler, electricity, total_water, feed_water, coal_carts, ash_carts, supply_temp, return_temp, outdoor_temp)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row["date"], row["boiler"], row["electricity"], row["total_water"],
                row["feed_water"], row["coal_carts"], row["ash_carts"],
                row["supply_temp"], row["return_temp"], row["outdoor_temp"]
            ))
        self.conn.commit()