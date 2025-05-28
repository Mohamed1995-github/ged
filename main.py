# main.py

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
        self.file_path = None
        self.setup_ui()
    
    def setup_ui(self):
        self.create_detailed_search()
        self.create_input_form()
        self.create_search_area()
        self.create_results_table()
    
    def create_detailed_search(self):
        frame = ttk.LabelFrame(self.root, text="Recherche détaillée")
        frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(frame, text="N° dossier (détaillé) :").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.detailed_search_entry = ttk.Entry(frame)
        self.detailed_search_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frame, text="Voir détails", command=self.show_document_details)\
            .grid(row=0, column=2, padx=5, pady=2)
    
    def create_input_form(self):
        input_frame = ttk.LabelFrame(self.root, text="Nouveau document")
        input_frame.pack(fill="x", padx=5, pady=5)
        
        labels = ["Numero du dossier", "Modele", "Langue", "Titer", "Date", "Departement", "Objet", "Contenu", "Type"]
        self.entries = {}
        for i, label in enumerate(labels):
            key = label.lower().replace(" ", "_")
            ttk.Label(input_frame, text=label).grid(row=i, column=0, padx=5, pady=2, sticky="w")
            if label == "Modele":
                widget = ttk.Combobox(input_frame, values=["Document logistlatif", "couriel"])
            elif label == "Langue":
                widget = ttk.Combobox(input_frame, values=["Français", "Arabe"])
            elif label == "Contenu":
                widget = tk.Text(input_frame, height=3)
            elif label == "Type":
                widget = ttk.Combobox(input_frame, values=["arrété", "décret", "projet de loi", "loi", "Courriel depart", "Couriel Arriver"])
            else:
                widget = ttk.Entry(input_frame)
            widget.grid(row=i, column=1, padx=5, pady=2, sticky="ew")
            self.entries[key] = widget
        
        # Emplacement physique
        emplacement_frame = ttk.LabelFrame(input_frame, text="Emplacement physique")
        emplacement_frame.grid(row=len(labels), column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        mapping = {
            "N° boîte d'archive": "numero_boite",
            "Salle": "salle",
            "Étagère": "etagere",
            "Rayon": "rayon",
        }
        self.emplacement_entries = {}
        for idx, (label, key) in enumerate(mapping.items()):
            ttk.Label(emplacement_frame, text=label).grid(row=idx//2, column=(idx%2)*2, padx=5, pady=2, sticky="w")
            ent = ttk.Entry(emplacement_frame)
            ent.grid(row=idx//2, column=(idx%2)*2+1, padx=5, pady=2, sticky="ew")
            self.emplacement_entries[key] = ent
        
        # Import de fichier + bouton Ajouter
        row = len(labels) + 1
        self.file_label = ttk.Label(input_frame, text="Aucun fichier sélectionné")
        self.file_label.grid(row=row, column=0, columnspan=2, pady=5)
        ttk.Button(input_frame, text="Importer un fichier", command=self.import_file)\
            .grid(row=row+1, column=0, columnspan=2, pady=5)
        ttk.Button(input_frame, text="Ajouter", command=self.add_document)\
            .grid(row=row+2, column=0, columnspan=2, pady=10)
    
    def import_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("PDF et images", "*.pdf;*.png;*.jpg;*.jpeg"), ("Tous", "*.*")]
        )
        if path:
            self.file_path = path
            self.file_label.config(text=os.path.basename(path))
    
    def create_search_area(self):
        frame = ttk.LabelFrame(self.root, text="Recherche")
        frame.pack(fill="x", padx=5, pady=5)
        fields = ["Numero_du_dossier", "Date", "Departement", "Type", "Objet"]
        self.search_entries = {}
        for i, fld in enumerate(fields):
            lbl = fld.replace("_", " ")
            ttk.Label(frame, text=lbl).grid(row=i//2, column=(i%2)*2, padx=5, pady=2, sticky="w")
            ent = ttk.Entry(frame)
            ent.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=2, sticky="ew")
            self.search_entries[fld.lower()] = ent
        
        # Recherche par emplacement
        em_frame = ttk.LabelFrame(frame, text="Recherche par emplacement")
        em_frame.grid(row=len(fields)//2+1, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        mapping = {
            "N° boîte d'archive": "numero_boite",
            "Salle": "salle",
            "Étagère": "etagere",
            "Rayon": "rayon",
        }
        self.emplacement_search_entries = {}
        for i, (label, key) in enumerate(mapping.items()):
            ttk.Label(em_frame, text=label).grid(row=i//2, column=(i%2)*2, padx=5, pady=2, sticky="w")
            ent = ttk.Entry(em_frame)
            ent.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=2, sticky="ew")
            self.emplacement_search_entries[key] = ent
        
        ttk.Button(frame, text="Rechercher", command=self.search_documents).grid(
            row=len(fields)//2+3, column=0, columnspan=2, pady=10)
        ttk.Button(frame, text="Réinitialiser", command=self.reset_search).grid(
            row=len(fields)//2+3, column=2, columnspan=2, pady=10)
    
    def create_results_table(self):
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=5, pady=5)
        ttk.Button(frame, text="Exporter les résultats", command=self.export_results)\
            .pack(side="top", anchor="ne", padx=5, pady=5)
        
        cols = ("numero_dossier","date","departement","type","objet",
                "numero_boite","salle","etagere","rayon","file_path")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings")
        headings = {
            "numero_dossier":"N° Dossier","date":"Date","departement":"Département",
            "type":"Type","objet":"Objet","numero_boite":"N° Boîte","salle":"Salle",
            "etagere":"Étagère","rayon":"Rayon","file_path":"Fichier"
        }
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=80 if c!="objet" and c!="file_path" else 150)
        self.tree.bind("<Double-1>", lambda e: self.open_selected_file())
        
        sb_y = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        sb_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
    
    def add_document(self):
        data = {
            'numero_dossier': self.entries['numero_du_dossier'].get().strip(),
            'date': self.entries['date'].get().strip(),
            'departement': self.entries['departement'].get().strip(),
            'type': self.entries['type'].get().strip() if 'type' in self.entries else '',
            'objet': self.entries['objet'].get().strip(),
            'numero_boite': self.emplacement_entries['numero_boite'].get().strip(),
            'salle': self.emplacement_entries['salle'].get().strip(),
            'etagere': self.emplacement_entries['etagere'].get().strip(),
            'rayon': self.emplacement_entries['rayon'].get().strip(),
            'file_path': self.file_path
        }
        # champs optionnels
        for opt in ('modele','langue','titer','contenu'):
            if opt in self.entries:
                if opt == 'contenu':
                    data[opt] = self.entries[opt].get("1.0", tk.END).strip()
                else:
                    data[opt] = self.entries[opt].get().strip()
        
        if not data['numero_dossier'] or not data['objet']:
            messagebox.showerror("Erreur", "Le numéro du dossier et l'objet sont obligatoires.")
            return
        
        try:
            self.db.add_document(**data)
            messagebox.showinfo("Succès", "Document ajouté !")
            # reset form
            for w in self.entries.values():
                if isinstance(w, tk.Text):
                    w.delete("1.0", tk.END)
                else:
                    w.delete(0, tk.END)
            for w in self.emplacement_entries.values():
                w.delete(0, tk.END)
            self.file_label.config(text="Aucun fichier sélectionné")
            self.file_path = None
            self.search_documents()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter :\n{e}")
    
    def reset_search(self):
        for e in (*self.search_entries.values(), *self.emplacement_search_entries.values()):
            e.delete(0, tk.END)
        self.search_documents()
    
    def search_documents(self):
        crit = {}
        for k, ent in self.search_entries.items():
            v = ent.get().strip()
            if v: crit[k] = v
        for k, ent in self.emplacement_search_entries.items():
            v = ent.get().strip()
            if v: crit[k] = v
        
        rows = self.db.search_documents(crit)
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for row in rows:
            # row = (id, numero_dossier, modele, langue, titer, date,
            #        departement, emplacement, numero_boite, salle, etagere,
            #        rayon, objet, contenu, type, file_path)
            vals = [
                row[1], row[5], row[6], row[14], row[12],
                row[8], row[9], row[10], row[11], row[15]
            ]
            self.tree.insert("", "end", values=vals)
        cnt = len(rows)
        messagebox.showinfo("Recherche", f"{cnt} document(s) trouvé(s).")
    
    def open_selected_file(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Ouvrir", "Sélectionnez d'abord un document.")
            return
        path = self.tree.item(sel[0], "values")[-1]
        if path and os.path.exists(path):
            webbrowser.open(path)
        else:
            messagebox.showerror("Erreur", "Fichier introuvable.")
    
    def export_results(self):
        if not self.tree.get_children():
            messagebox.showinfo("Export", "Aucun résultat à exporter.")
            return
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = f"export_{now}.csv"
        fp = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv")],
                                          initialfile=default)
        if not fp: return
        headers = [self.tree.heading(c)["text"] for c in self.tree["columns"]]
        rows = [self.tree.item(i,"values") for i in self.tree.get_children()]
        try:
            with open(fp, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(headers)
                w.writerows(rows)
            messagebox.showinfo("Export", f"Exporté dans {fp}")
            if messagebox.askyesno("Ouvrir", "Ouvrir le fichier exporté ?"):
                webbrowser.open(fp)
        except Exception as e:
            messagebox.showerror("Erreur d'export", str(e))
    
    def show_document_details(self):
        numero = self.detailed_search_entry.get().strip()
        if not numero:
            messagebox.showwarning("Recherche détaillée", "Entrez un numéro de dossier.")
            return
        doc = self.db.get_document_by_numero(numero)
        if not doc:
            messagebox.showinfo("Recherche détaillée", f"Aucun document pour {numero}.")
            return
        
        popup = tk.Toplevel(self.root)
        popup.title(f"Détails du dossier {numero}")
        popup.geometry("+200+200")
        champs = [
            ("N° dossier",     doc.get("numero_dossier")),
            ("Modèle",         doc.get("modele")),
            ("Langue",         doc.get("langue")),
            ("Titre",          doc.get("titer")),
            ("Date",           doc.get("date")),
            ("Département",    doc.get("departement")),
            ("Objet",          doc.get("objet")),
            ("Contenu",        doc.get("contenu")),
            ("Type",           doc.get("type")),
            ("N° boîte",       doc.get("numero_boite")),
            ("Salle",          doc.get("salle")),
            ("Étagère",        doc.get("etagere")),
            ("Rayon",          doc.get("rayon")),
            ("Chemin fichier", doc.get("file_path")),
        ]
        for i, (lbl, val) in enumerate(champs):
            ttk.Label(popup, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=5, pady=2)
            ttk.Label(popup, text=val or "").grid(row=i, column=1, sticky="w", padx=5, pady=2)
        
        fp = doc.get("file_path")
        if fp:
            def _open():
                if os.path.exists(fp):
                    webbrowser.open(fp)
                else:
                    messagebox.showerror("Ouvrir", "Fichier introuvable.")
            ttk.Button(popup, text="Ouvrir le fichier", command=_open)\
                .grid(row=len(champs), column=0, columnspan=2, pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = GEDApp(root)
    root.mainloop()
