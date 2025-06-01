import sqlite3
import os

class Database:
    def __init__(self, db_path='documents.db'):
        self.conn = sqlite3.connect(db_path)
        self.create_tables()
    
    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_dossier TEXT NOT NULL,
            modele         TEXT,
            langue         TEXT,
            titer          TEXT,
            date           TEXT,
            departement    TEXT,
            emplacement    TEXT,
            numero_boite   TEXT,
            salle          TEXT,
            etagere        TEXT,
            rayon          TEXT,
            objet          TEXT NOT NULL,
            numero_piece   TEXT,
            type           TEXT,
            file_path      TEXT,
            statut_archivage TEXT
        )
        ''')
        self.conn.commit()
    
    def add_document(self, **kwargs):
        cursor = self.conn.cursor()
        fields       = ', '.join(kwargs.keys())
        placeholders = ', '.join('?' for _ in kwargs)
        values       = tuple(kwargs.values())
        query = f'INSERT INTO documents ({fields}) VALUES ({placeholders})'
        cursor.execute(query, values)
        self.conn.commit()
    
    def search_documents(self, criteria):
        cursor = self.conn.cursor()
        conds  = []
        vals   = []
        for field, value in criteria.items():
            if value:
                conds.append(f"{field} LIKE ?")
                vals.append(f"%{value}%")
        if conds:
            q = "SELECT * FROM documents WHERE " + " AND ".join(conds)
            cursor.execute(q, vals)
        else:
            cursor.execute("SELECT * FROM documents")
        return cursor.fetchall()
    
    def search_by_numero_dossier_or_piece(self, search_term):
        """Recherche par numéro de dossier OU numéro de pièce"""
        cursor = self.conn.cursor()
        query = """
        SELECT * FROM documents 
        WHERE (numero_dossier IS NOT NULL AND numero_dossier LIKE ?) 
           OR (numero_piece IS NOT NULL AND numero_piece LIKE ?)
        """
        search_pattern = f"%{search_term}%"
        cursor.execute(query, (search_pattern, search_pattern))
        return cursor.fetchall()
    
    def get_documents_by_status(self, statut):
        """Récupère tous les documents ayant un statut d'archivage donné"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM documents WHERE statut_archivage = ?", (statut,))
        return cursor.fetchall()
    
    def get_document_by_id(self, doc_id):
        """Renvoie un dict colonne→valeur ou None pour un ID donné."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM documents WHERE id = ?", (doc_id,))
        row = cursor.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cursor.description]
        return dict(zip(cols, row))
    
    def get_document_by_numero_piece(self, numero_piece):
        """Renvoie un dict colonne→valeur ou None."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM documents WHERE numero_piece = ?", (numero_piece,))
        row = cursor.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cursor.description]
        return dict(zip(cols, row))
    
    def update_document(self, document_id, **kwargs):
        cursor = self.conn.cursor()
        parts = []
        vals  = []
        for f, v in kwargs.items():
            parts.append(f"{f} = ?")
            vals.append(v)
        vals.append(document_id)
        q = f"UPDATE documents SET {', '.join(parts)} WHERE id = ?"
        cursor.execute(q, vals)
        self.conn.commit()
    
    def delete_document(self, document_id):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        self.conn.commit()
    
    def get_all_statuts(self):
        """Récupère tous les statuts d'archivage existants"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT statut_archivage FROM documents WHERE statut_archivage IS NOT NULL AND statut_archivage != ''")
        return [row[0] for row in cursor.fetchall()]
    
    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()