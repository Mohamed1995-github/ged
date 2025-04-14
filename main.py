import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import webbrowser
import csv
from datetime import datetime
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
        
        labels = ["Numero du dossier", "Modele", "Langue", "Titer","Date", "Departement", "Emplacement physique", "Objet", "Contenu", "Type"]
        self.entries = {}
        
        for i, label in enumerate(labels):
            ttk.Label(input_frame, text=label).grid(row=i, column=0, padx=5, pady=2)
            if label == "Modele":
                self.entries[label.lower()] = ttk.Combobox(input_frame, values=["Document logistlatif", "couriel"])
            elif label == "Langue":
                self.entries[label.lower()] = ttk.Combobox(input_frame, values=["Français", "Arabe"])
            elif label == "Contenu":
                self.entries[label.lower()] = tk.Text(input_frame, height=3)
            elif label == "Type":
                self.entries[label.lower()] = ttk.Combobox(input_frame, values=["arrété", "décret", "projet de loi", "loi","Courriel depart", "Couriel Arriver"])
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
        
        # Créer des champs de recherche
        search_fields = ["Numero du dossier", "Date", "Departement", "Type", "Objet"]
        self.search_entries = {}
        
        for i, field in enumerate(search_fields):
            ttk.Label(search_frame, text=field).grid(row=i//2, column=(i%2)*2, padx=5, pady=2)
            self.search_entries[field.lower()] = ttk.Entry(search_frame)
            self.search_entries[field.lower()].grid(row=i//2, column=(i%2)*2+1, padx=5, pady=2)

        # Bouton de recherche avec style amélioré
        search_button = ttk.Button(search_frame, text="Rechercher", command=self.search_documents)
        search_button.grid(row=len(search_fields)//2 + 1, column=0, columnspan=2, pady=10)
        
        # Bouton de réinitialisation des filtres
        reset_button = ttk.Button(search_frame, text="Réinitialiser", command=self.reset_search)
        reset_button.grid(row=len(search_fields)//2 + 1, column=2, columnspan=2, pady=10)
        
        # Ajouter un style au bouton de recherche pour le rendre plus visible
        style = ttk.Style()
        style.configure("Search.TButton", foreground="white", background="#4CAF50", font=('Helvetica', 10, 'bold'))
        search_button.configure(style="Search.TButton")
    
    def reset_search(self):
        # Réinitialiser tous les champs de recherche
        for entry in self.search_entries.values():
            entry.delete(0, tk.END)
        
        # Rafraîchir la liste pour afficher tous les documents
        self.search_documents()
    
    def create_results_table(self):
        # Créer un cadre pour le tableau et les boutons d'action
        results_frame = ttk.Frame(self.root)
        results_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Bouton d'exportation au-dessus du tableau
        export_button = ttk.Button(results_frame, text="Exporter les résultats", command=self.export_results)
        export_button.pack(side="top", anchor="ne", padx=5, pady=5)
        
        # Style pour le bouton d'exportation
        style = ttk.Style()
        style.configure("Export.TButton", foreground="white", background="#007BFF", font=('Helvetica', 10, 'bold'))
        export_button.configure(style="Export.TButton")
        
        # Tableau des résultats
        columns = ("numero_dossier", "date", "departement", "type", "objet", "emplacement", "file_path")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        for col in columns[:-1]:  # On n'affiche pas le chemin du fichier
            self.tree.heading(col, text=col.title())
            self.tree.column(col, width=100)
        
        self.tree.heading("file_path", text="Fichier")
        self.tree.column("file_path", width=120)
        
        # Ajouter un menu contextuel pour ouvrir le fichier
        self.tree.bind("<Double-1>", lambda event: self.open_selected_file())
        
        # Ajouter des barres de défilement
        scrollbar_y = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

    def add_document(self):
        doc_data = {
            'numero_dossier': self.entries['numero du dossier'].get(),
            'date': self.entries['date'].get(),
            'departement': self.entries['departement'].get(),
            'type': self.entries['type'].get(),
            'objet': self.entries['objet'].get(),
            'emplacement': self.entries['emplacement physique'].get(),
            'file_path': self.file_path
        }
        
        # Optional fields not shown in table but stored in DB
        if 'modele' in self.entries:
            doc_data['modele'] = self.entries['modele'].get()
        if 'langue' in self.entries:
            doc_data['langue'] = self.entries['langue'].get()
        if 'titer' in self.entries:
            doc_data['titer'] = self.entries['titer'].get()
        if 'contenu' in self.entries:
            doc_data['contenu'] = self.entries['contenu'].get("1.0", tk.END).strip()
        
        if not doc_data['numero_dossier'] or not doc_data['objet']:
            messagebox.showerror("Erreur", "Le numéro du dossier et l'objet sont obligatoires.")
            return
        
        try:
            self.db.add_document(**doc_data)
            messagebox.showinfo("Succès", "Document ajouté avec succès!")
            
            # Réinitialiser les champs du formulaire
            for entry in self.entries.values():
                if isinstance(entry, tk.Text):
                    entry.delete("1.0", tk.END)
                else:
                    entry.delete(0, tk.END)
                    
            self.file_label.config(text="Aucun fichier sélectionné")
            self.file_path = None
            self.search_documents()  # Rafraîchir la liste
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter le document: {str(e)}")
    
    def search_documents(self):
        criteria = {}
        
        # Récupérer les valeurs de recherche non vides
        for field, entry in self.search_entries.items():
            value = entry.get().strip()
            if value:  # Ne pas inclure les champs vides
                # Convertir les noms des champs pour correspondre à la BDD
                db_field = field.replace(" ", "_") if " " in field else field
                criteria[db_field] = value
        
        results = self.db.search_documents(criteria)
        
        # Effacer les résultats précédents
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Afficher les nouveaux résultats
        for doc in results:
            values_to_display = [
                doc[1],  # numero_dossier
                doc[5],  # date
                doc[6],  # departement
                doc[10],  # type
                doc[8],  # objet
                doc[7],  # emplacement
                doc[11]   # file_path
            ]
            self.tree.insert("", "end", values=values_to_display)
        
        # Informer l'utilisateur du nombre de résultats
        if len(results) == 0:
            messagebox.showinfo("Recherche", "Aucun document trouvé.")
        else:
            messagebox.showinfo("Recherche", f"{len(results)} document(s) trouvé(s).")

    def open_selected_file(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Aucun fichier", "Veuillez sélectionner un document contenant un fichier.")
            return
        
        file_path = self.tree.item(selected_item[0], "values")[-1]
        if file_path and file_path != "None" and os.path.exists(file_path):
            webbrowser.open(file_path)  # Ouvre le fichier avec l'application par défaut
        else:
            messagebox.showerror("Erreur", "Fichier introuvable.")
            
    def export_results(self):
        """Exporte les résultats de la recherche dans un fichier CSV"""
        if not self.tree.get_children():
            messagebox.showinfo("Export", "Aucun résultat à exporter.")
            return
            
        # Demander à l'utilisateur où sauvegarder le fichier
        current_date = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"export_documents_{current_date}.csv"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")],
            initialfile=default_filename
        )
        
        if not file_path:
            return  # L'utilisateur a annulé
            
        try:
            # Récupérer les en-têtes
            headers = [self.tree.heading(column)["text"] for column in self.tree.cget("columns")]
            
            # Récupérer toutes les lignes
            rows = []
            for item_id in self.tree.get_children():
                values = self.tree.item(item_id, "values")
                rows.append(values)
                
            # Écrire dans le fichier CSV
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=';')
                writer.writerow(headers)
                writer.writerows(rows)
                
            messagebox.showinfo("Export réussi", f"Les données ont été exportées avec succès dans:\n{file_path}")
            
            # Proposer d'ouvrir le fichier
            if messagebox.askyesno("Ouvrir le fichier", "Voulez-vous ouvrir le fichier exporté?"):
                webbrowser.open(file_path)
                
        except Exception as e:
            messagebox.showerror("Erreur d'exportation", f"Une erreur s'est produite lors de l'exportation:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GEDApp(root)
    root.mainloop()