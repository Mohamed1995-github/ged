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
    
    def get_document(self, document_id):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM documents WHERE id = ?", (document_id,))
        return cursor.fetchone()
    
    def update_document(self, document_id, **kwargs):
        cursor = self.conn.cursor()
        
        # Construire les parties SET de la requête
        set_parts = []
        values = []
        
        for field, value in kwargs.items():
            set_parts.append(f"{field} = ?")
            values.append(value)
        
        # Ajouter l'ID à la fin des valeurs pour le WHERE
        values.append(document_id)
        
        query = f"UPDATE documents SET {', '.join(set_parts)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()
    
    def delete_document(self, document_id):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        self.conn.commit()
    
    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()