import sqlite3

# Function to initialize database and store complaints
def init_db():
    conn = sqlite3.connect('complaints.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            observation TEXT,
            priority TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Function to save new complaints to the database
def save_complaint(observation, priority):
    conn = sqlite3.connect('complaints.db')
    cursor = conn.cursor()
    cursor.execute('INSERT INTO complaints (observation, priority) VALUES (?, ?)', (observation, priority))
    conn.commit()
    conn.close()

# Call this once to initialize the database
init_db()
