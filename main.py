import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import webbrowser
from database import Database

class GEDApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestion Électronique des Documents")
        self.db = Database()
        self.file_path = None  # Pour stocker le chemin du fichier sélectionné
        
        self.setup_ui()
    
    def setup_ui(self):
        self.create_input_form()
        self.create_search_area()
        self.create_results_table()
    
    def create_input_form(self):
        input_frame = ttk.LabelFrame(self.root, text="Nouveau document")
        input_frame.pack(fill="x", padx=5, pady=5)
        
        labels = ["Référence", "Date", "Expéditeur", "Destinataire", "Objet", "Contenu", "Type"]
        self.entries = {}
        
        for i, label in enumerate(labels):
            ttk.Label(input_frame, text=label).grid(row=i, column=0, padx=5, pady=2)
            if label == "Contenu":
                self.entries[label.lower()] = tk.Text(input_frame, height=3)
            elif label == "Type":
                self.entries[label.lower()] = ttk.Combobox(input_frame, values=["Manuscrit", "Courriel"])
            else:
                self.entries[label.lower()] = ttk.Entry(input_frame)
            self.entries[label.lower()].grid(row=i, column=1, padx=5, pady=2)
        
        # Bouton pour importer un fichier
        self.file_label = ttk.Label(input_frame, text="Aucun fichier sélectionné")
        self.file_label.grid(row=len(labels), column=0, columnspan=2, pady=5)
        
        ttk.Button(input_frame, text="Importer un fichier", command=self.import_file).grid(
            row=len(labels) + 1, column=0, columnspan=2, pady=5
        )

        ttk.Button(input_frame, text="Ajouter", command=self.add_document).grid(
            row=len(labels) + 2, column=0, columnspan=2, pady=10
        )

    def import_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Fichiers PDF et images", "*.pdf;*.png;*.jpg;*.jpeg"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))  # Afficher le nom du fichier sélectionné
    
    def create_search_area(self):
        search_frame = ttk.LabelFrame(self.root, text="Recherche")
        search_frame.pack(fill="x", padx=5, pady=5)
        
        self.search_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Rechercher", command=self.search_documents).pack(side="left", padx=5)
    
    def create_results_table(self):
        columns = ("reference", "date", "expediteur", "destinataire", "objet", "type_doc", "file_path")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        
        for col in columns[:-1]:  # On n'affiche pas le chemin du fichier
            self.tree.heading(col, text=col.title())
            self.tree.column(col, width=100)
        
        self.tree.heading("file_path", text="Fichier", command=self.open_selected_file)
        self.tree.column("file_path", width=120)
        
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

    def add_document(self):
        doc_data = {
            'reference': self.entries['référence'].get(),
            'date': self.entries['date'].get(),
            'expediteur': self.entries['expéditeur'].get(),
            'destinataire': self.entries['destinataire'].get(),
            'objet': self.entries['objet'].get(),
            'contenu': self.entries['contenu'].get("1.0", tk.END).strip(),
            'type_doc': self.entries['type'].get(),
            'file_path': self.file_path
        }
        
        if not doc_data['reference'] or not doc_data['objet']:
            messagebox.showerror("Erreur", "La référence et l'objet sont obligatoires.")
            return
        
        self.db.add_document(**doc_data)
        self.file_label.config(text="Aucun fichier sélectionné")
        self.file_path = None
        self.search_documents()  # Rafraîchir la liste
    
    def search_documents(self):
        criteria = {'objet': self.search_var.get()}
        results = self.db.search_documents(criteria)
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for doc in results:
            self.tree.insert("", "end", values=doc[1:])  # On affiche les résultats sans l'ID

    def open_selected_file(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Aucun fichier", "Veuillez sélectionner un document contenant un fichier.")
            return
        
        file_path = self.tree.item(selected_item[0], "values")[-1]
        if file_path and os.path.exists(file_path):
            webbrowser.open(file_path)  # Ouvre le fichier avec l'application par défaut
        else:
            messagebox.showerror("Erreur", "Fichier introuvable.")

if __name__ == "__main__":
    root = tk.Tk()
    app = GEDApp(root)
    root.mainloop()
