import sqlite3
import os

class Database:
    def __init__(self):
        self.conn = sqlite3.connect('documents.db')
        self.create_tables()
    
    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_dossier TEXT NOT NULL,
            modele TEXT,
            langue TEXT,
            titer TEXT,
            date TEXT,
            departement TEXT,
            emplacement TEXT,
            objet TEXT NOT NULL,
            contenu TEXT,
            type TEXT,
            file_path TEXT
        )
        ''')
        self.conn.commit()
    
    def add_document(self, **kwargs):
        cursor = self.conn.cursor()
        fields = ', '.join(kwargs.keys())
        placeholders = ', '.join(['?' for _ in kwargs])
        values = tuple(kwargs.values())
        
        query = f'INSERT INTO documents ({fields}) VALUES ({placeholders})'
        cursor.execute(query, values)
        self.conn.commit()
    
    def search_documents(self, criteria):
        cursor = self.conn.cursor()
        conditions = []
        values = []
        
        for field, value in criteria.items():
            if value:
                conditions.append(f"{field} LIKE ?")
                values.append(f"%{value}%")
        
        if conditions:
            query = "SELECT * FROM documents WHERE " + " AND ".join(conditions)
            cursor.execute(query, values)
        else:
            cursor.execute("SELECT * FROM documents")
            
        return cursor.fetchall()
