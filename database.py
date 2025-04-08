import sqlite3
import os

class Database:
    def __init__(self):
        self.conn = sqlite3.connect('ged.db')
        self.create_tables()
    
    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reference TEXT UNIQUE,
                date TEXT,
                expediteur TEXT,
                destinataire TEXT,
                objet TEXT,
                contenu TEXT,
                type_doc TEXT,
                file_path TEXT
            )
        ''')
        self.conn.commit()
    
    def add_document(self, reference, date, expediteur, destinataire, objet, contenu, type_doc, file_path=None):
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO documents 
            (reference, date, expediteur, destinataire, objet, contenu, type_doc, file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (reference, date, expediteur, destinataire, objet, contenu, type_doc, file_path))
        self.conn.commit()
    
    def search_documents(self, criteria):
        cursor = self.conn.cursor()
        query = "SELECT * FROM documents WHERE 1=1"
        params = []
        
        for key, value in criteria.items():
            if value:
                query += f" AND {key} LIKE ?"
                params.append(f"%{value}%")
        
        cursor.execute(query, params)
        return cursor.fetchall()
