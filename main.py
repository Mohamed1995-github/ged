# main.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, webbrowser, csv
from datetime import datetime
from database import Database

class GEDApp:
    def __init__(self, root):
        self.root      = root
        self.root.title("Gestion Électronique des Documents")
        self.db        = Database()
        self.file_path = None
        self.current_document_id = None  # Pour la modification
        self.setup_ui()
    
    def setup_ui(self):
        self.create_status_buttons()
        self.create_dossier_search()  # Recherche combinée par numéro de dossier/pièce
        self.create_detailed_search()
        self.create_input_form()
        self.create_results_table()
    
    def create_status_buttons(self):
        """Crée les boutons de statut d'archivage"""
        frame = ttk.LabelFrame(self.root, text="Filtrer par statut d'archivage")
        frame.pack(fill="x", padx=5, pady=5)
        
        # Boutons prédéfinis
        statuts = [
            "archifage courants",
            "archifage intermediers", 
            "archifage Finals"
        ]
        
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=5)
        
        for i, statut in enumerate(statuts):
            btn = ttk.Button(button_frame, text=statut, 
                           command=lambda s=statut: self.filter_by_status(s))
            btn.grid(row=0, column=i, padx=5, pady=2)
        
        # Bouton pour afficher tous les documents
        ttk.Button(button_frame, text="Tous les documents", 
                  command=self.show_all_documents).grid(row=0, column=len(statuts), padx=5, pady=2)
    
    def create_dossier_search(self):
        """Crée la section de recherche par numéro de dossier ou de pièce"""
        frame = ttk.LabelFrame(self.root, text="Recherche par N° dossier ou N° pièce")
        frame.pack(fill="x", padx=5, pady=5)
        
        # Frame pour organiser les éléments horizontalement
        search_frame = ttk.Frame(frame)
        search_frame.pack(pady=5)
        
        ttk.Label(search_frame, text="N° dossier/pièce :").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.dossier_search_entry = ttk.Entry(search_frame, width=30)
        self.dossier_search_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Button(search_frame, text="Rechercher", 
                  command=self.search_by_dossier_or_piece).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(search_frame, text="Réinitialiser", 
                  command=self.reset_dossier_search).grid(row=0, column=3, padx=5, pady=2)
    
    def create_detailed_search(self):
        """Section de recherche détaillée pour afficher les détails complets"""
        f = ttk.LabelFrame(self.root, text="Recherche détaillée")
        f.pack(fill="x", padx=5, pady=5)
        
        # Frame pour organiser les éléments horizontalement
        search_frame = ttk.Frame(f)
        search_frame.pack(pady=5)
        
        # Recherche par numéro de pièce
        ttk.Label(search_frame, text="N° pièce :").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.detailed_search_entry = ttk.Entry(search_frame, width=25)
        self.detailed_search_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(search_frame, text="Voir détails", command=self.show_document_details)\
            .grid(row=0, column=2, padx=5, pady=2)
        
        # Recherche par numéro de dossier
        ttk.Label(search_frame, text="N° dossier :").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.detailed_dossier_entry = ttk.Entry(search_frame, width=25)
        self.detailed_dossier_entry.grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(search_frame, text="Voir détails", command=self.show_dossier_details)\
            .grid(row=1, column=2, padx=5, pady=2)
        
        # Recherche par objet
        ttk.Label(search_frame, text="Objet :").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.detailed_objet_entry = ttk.Entry(search_frame, width=25)
        self.detailed_objet_entry.grid(row=2, column=1, padx=5, pady=2)
        ttk.Button(search_frame, text="Rechercher objet", command=self.search_by_objet)\
            .grid(row=2, column=2, padx=5, pady=2)
        
    def create_input_form(self):
        frame = ttk.LabelFrame(self.root, text="Nouveau document")
        frame.pack(fill="x", padx=5, pady=5)
        
        # --- Champs standards ---
        labels = [
            "Numero du dossier", "Modele", "Langue",
            "Titer", "Date", "Departement",
            "Objet", "Numero de piece", "Type"
        ]
        self.entries = {}
        for i, label in enumerate(labels):
            key = label.lower().replace(" ", "_")
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            if label == "Modele":
                w = ttk.Combobox(frame, values=["Document logistlatif", "couriel"])
            elif label == "Langue":
                w = ttk.Combobox(frame, values=["Français", "Arabe"])
            elif label == "Type":
                w = ttk.Combobox(frame, values=[
                    "arrété","décret","projet de loi","loi",
                    "Courriel depart","Couriel Arriver"
                ])
            else:
                w = ttk.Entry(frame)
            w.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            self.entries[key] = w
        
        # --- Statut d'archivage ---
        ttk.Label(frame, text="Statut archivage").grid(
            row=len(labels), column=0, sticky="w", padx=5, pady=2)
        self.archivage_entry = ttk.Combobox(frame, values=[
            "archifage courants",
            "archifage intermediers",
            "archifage Finals"
        ])
        self.archivage_entry.grid(row=len(labels), column=1, sticky="ew", padx=5, pady=2)
        
        # --- Emplacement physique ---
        em = ttk.LabelFrame(frame, text="Emplacement physique")
        em.grid(row=len(labels)+1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        mapping = {
            "N° boîte d'archive": "numero_boite",
            "Salle":               "salle",
            "Étagère":             "etagere",
            "Rayon":               "rayon",
        }
        self.emplacement_entries = {}
        for idx, (lab, key) in enumerate(mapping.items()):
            ttk.Label(em, text=lab).grid(row=idx//2, column=(idx%2)*2, sticky="w", padx=5, pady=2)
            ent = ttk.Entry(em)
            ent.grid(row=idx//2, column=(idx%2)*2+1, sticky="ew", padx=5, pady=2)
            self.emplacement_entries[key] = ent
        
        # --- Fichier + boutons ---
        r = len(labels) + 2
        self.file_label = ttk.Label(frame, text="Aucun fichier sélectionné")
        self.file_label.grid(row=r, column=0, columnspan=2, pady=5)
        ttk.Button(frame, text="Importer un fichier", command=self.import_file).grid(
            row=r+1, column=0, columnspan=2, pady=5)
        
        # Frame pour les boutons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=r+2, column=0, columnspan=2, pady=10)
        
        self.add_btn = ttk.Button(btn_frame, text="Ajouter", command=self.add_document)
        self.add_btn.pack(side="left", padx=5)
        
        self.modify_btn = ttk.Button(btn_frame, text="Modifier", command=self.modify_document, state="disabled")
        self.modify_btn.pack(side="left", padx=5)
        
        self.cancel_btn = ttk.Button(btn_frame, text="Annuler", command=self.cancel_modification, state="disabled")
        self.cancel_btn.pack(side="left", padx=5)
    
    def import_file(self):
        p = filedialog.askopenfilename(
            filetypes=[("PDF et images","*.pdf;*.png;*.jpg;*.jpeg"), ("Tous","*.*")]
        )
        if p:
            self.file_path = p
            self.file_label.config(text=os.path.basename(p))
    
    def create_results_table(self):
        f = ttk.Frame(self.root)
        f.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Boutons d'action
        btn_frame = ttk.Frame(f)
        btn_frame.pack(side="top", fill="x", padx=5, pady=5)
        
        ttk.Button(btn_frame, text="Exporter les résultats", command=self.export_results)\
            .pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Modifier sélectionné", command=self.prepare_modification)\
            .pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Supprimer sélectionné", command=self.delete_selected)\
            .pack(side="right", padx=5)
        
        # Colonnes étendues pour plus de détails
        cols = (
            "id","numero_dossier","numero_piece","date","departement","type","objet",
            "modele","langue","titer","numero_boite","salle","etagere","rayon",
            "statut_archivage","file_path"
        )
        self.tree = ttk.Treeview(f, columns=cols, show="headings")
        hdrs = {
            "id":"ID",
            "numero_dossier":"N° Dossier",
            "numero_piece":"N° Pièce",
            "date":"Date",
            "departement":"Département",
            "type":"Type",
            "objet":"Objet",
            "modele":"Modèle",
            "langue":"Langue",
            "titer":"Titre",
            "numero_boite":"N° Boîte",
            "salle":"Salle",
            "etagere":"Étagère",
            "rayon":"Rayon",
            "statut_archivage":"Statut",
            "file_path":"Fichier"
        }
        
        for c in cols:
            self.tree.heading(c, text=hdrs[c])
            if c == "id":
                self.tree.column(c, width=40)
            elif c in ("numero_dossier", "numero_piece"):
                self.tree.column(c, width=90)
            elif c in ("objet", "titer"):
                self.tree.column(c, width=150)
            elif c == "file_path":
                self.tree.column(c, width=200)
            elif c in ("date", "departement", "type"):
                self.tree.column(c, width=100)
            else:
                self.tree.column(c, width=80)
        
        self.tree.bind("<Double-1>", lambda e: self.open_selected_file())
        
        sy = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(f, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        
        # Charger tous les documents au démarrage
        self.show_all_documents()
    
    def search_by_dossier_or_piece(self):
        """Recherche les documents par numéro de dossier OU numéro de pièce"""
        search_term = self.dossier_search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("Recherche", "Veuillez entrer un numéro de dossier ou de pièce.")
            return
        
        # Utiliser la nouvelle méthode de recherche combinée
        rows = self.db.search_by_numero_dossier_or_piece(search_term)
        self.populate_tree(rows)
        
        if rows:
            messagebox.showinfo("Recherche", f"{len(rows)} document(s) trouvé(s) pour '{search_term}'.")
        else:
            messagebox.showinfo("Recherche", f"Aucun document trouvé pour '{search_term}'.")
    
    def reset_dossier_search(self):
        """Remet à zéro la recherche et affiche tous les documents"""
        self.dossier_search_entry.delete(0, tk.END)
        self.show_all_documents()
    
    def filter_by_status(self, statut):
        """Affiche les documents filtrés par statut"""
        rows = self.db.get_documents_by_status(statut)
        self.populate_tree(rows)
        messagebox.showinfo("Filtre", f"{len(rows)} document(s) avec le statut '{statut}'.")
    
    def show_all_documents(self):
        """Affiche tous les documents"""
        rows = self.db.search_documents({})
        self.populate_tree(rows)
    
    def populate_tree(self, rows):
        """Remplit le tableau avec les données détaillées"""
        # Vider d'abord le tableau
        for iid in self.tree.get_children(): 
            self.tree.delete(iid)
        
        for row in rows:
            # Structure réelle de la base : 18 colonnes
            # (id, numero_dossier, modele, langue, titer, date, departement, 
            #  emplacement, objet, contenu, type, file_path, numero_boite, 
            #  salle, etagere, rayon, numero_piece, statut_archivage)
            
            vals = [
                str(row[0]) if row[0] is not None else "",     # id
                str(row[1]) if row[1] is not None else "",     # numero_dossier
                str(row[16]) if row[16] is not None else "",   # numero_piece
                str(row[5]) if row[5] is not None else "",     # date
                str(row[6]) if row[6] is not None else "",     # departement
                str(row[10]) if row[10] is not None else "",   # type
                str(row[8]) if row[8] is not None else "",     # objet
                str(row[2]) if row[2] is not None else "",     # modele
                str(row[3]) if row[3] is not None else "",     # langue
                str(row[4]) if row[4] is not None else "",     # titer
                str(row[12]) if row[12] is not None else "",   # numero_boite
                str(row[13]) if row[13] is not None else "",   # salle
                str(row[14]) if row[14] is not None else "",   # etagere
                str(row[15]) if row[15] is not None else "",   # rayon
                str(row[17]) if row[17] is not None else "",   # statut_archivage
                str(row[11]) if row[11] is not None else ""    # file_path
            ]
            
            self.tree.insert("", "end", values=vals)
        
        # Forcer le rafraîchissement du tableau
        self.tree.update_idletasks()
    
    def prepare_modification(self):
        """Prépare la modification d'un document sélectionné"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Modification", "Sélectionnez d'abord un document.")
            return
        
        # Récupérer l'ID du document (première colonne)
        doc_id = self.tree.item(sel[0], "values")[0]
        doc = self.db.get_document_by_id(doc_id)
        
        if not doc:
            messagebox.showerror("Erreur", "Document introuvable.")
            return
        
        # Remplir le formulaire avec les données existantes
        self.current_document_id = doc_id
        self.fill_form_with_document(doc)
        
        # Changer l'état des boutons
        self.add_btn.config(state="disabled")
        self.modify_btn.config(state="normal")
        self.cancel_btn.config(state="normal")
        
        messagebox.showinfo("Modification", "Document chargé pour modification. Modifiez les champs puis cliquez sur 'Modifier'.")
    
    def fill_form_with_document(self, doc):
        """Remplit le formulaire avec les données du document"""
        # Nettoyer d'abord
        self.clear_form()
        
        # Remplir les champs
        mapping = {
            'numero_du_dossier': doc.get('numero_dossier', ''),
            'modele': doc.get('modele', ''),
            'langue': doc.get('langue', ''),
            'titer': doc.get('titer', ''),
            'date': doc.get('date', ''),
            'departement': doc.get('departement', ''),
            'objet': doc.get('objet', ''),
            'numero_de_piece': doc.get('numero_piece', ''),
            'type': doc.get('type', '')
        }
        
        for key, value in mapping.items():
            if key in self.entries:
                widget = self.entries[key]
                if isinstance(widget, ttk.Combobox):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, value)
        
        # Statut d'archivage
        self.archivage_entry.set(doc.get('statut_archivage', ''))
        
        # Emplacement
        emplacement_mapping = {
            'numero_boite': doc.get('numero_boite', ''),
            'salle': doc.get('salle', ''),
            'etagere': doc.get('etagere', ''),
            'rayon': doc.get('rayon', '')
        }
        
        for key, value in emplacement_mapping.items():
            if key in self.emplacement_entries:
                self.emplacement_entries[key].delete(0, tk.END)
                self.emplacement_entries[key].insert(0, value)
        
        # Fichier
        file_path = doc.get('file_path', '')
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
        else:
            self.file_path = None
            self.file_label.config(text="Aucun fichier sélectionné")
    
    def modify_document(self):
        """Modifie le document en cours"""
        if not self.current_document_id:
            return
        
        data = {
            'numero_dossier': self.entries['numero_du_dossier'].get().strip(),
            'modele'        : self.entries['modele'].get().strip(),
            'langue'        : self.entries['langue'].get().strip(),
            'titer'         : self.entries['titer'].get().strip(),
            'date'          : self.entries['date'].get().strip(),
            'departement'   : self.entries['departement'].get().strip(),
            'objet'         : self.entries['objet'].get().strip(),
            'numero_piece'  : self.entries['numero_de_piece'].get().strip(),
            'type'          : self.entries['type'].get().strip(),
            'statut_archivage': self.archivage_entry.get().strip(),
            'numero_boite'  : self.emplacement_entries['numero_boite'].get().strip(),
            'salle'         : self.emplacement_entries['salle'].get().strip(),
            'etagere'       : self.emplacement_entries['etagere'].get().strip(),
            'rayon'         : self.emplacement_entries['rayon'].get().strip(),
            'file_path'     : self.file_path
        }
        
        # Validations
        if not data['numero_piece'] or not data['objet']:
            messagebox.showerror("Erreur", "Le N° de pièce et l'objet sont obligatoires.")
            return
        
        try:
            self.db.update_document(self.current_document_id, **data)
            messagebox.showinfo("Succès", "Document modifié avec succès !")
            self.cancel_modification()
            self.show_all_documents()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de modifier :\n{e}")
    
    def cancel_modification(self):
        """Annule la modification en cours"""
        self.current_document_id = None
        self.clear_form()
        
        # Remettre les boutons dans leur état normal
        self.add_btn.config(state="normal")
        self.modify_btn.config(state="disabled")
        self.cancel_btn.config(state="disabled")
    
    def delete_selected(self):
        """Supprime le document sélectionné"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Suppression", "Sélectionnez d'abord un document.")
            return
        
        values = self.tree.item(sel[0], "values")
        doc_id = values[0]  # ID
        numero_piece = values[2] if values[2] else "Sans numéro"  # N° Pièce
        
        if messagebox.askyesno("Confirmation", f"Êtes-vous sûr de vouloir supprimer le document N° {numero_piece} ?"):
            try:
                self.db.delete_document(doc_id)
                messagebox.showinfo("Succès", "Document supprimé !")
                self.show_all_documents()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de supprimer :\n{e}")
    
    def clear_form(self):
        """Vide le formulaire"""
        for w in self.entries.values():
            if isinstance(w, ttk.Combobox):
                w.set('')
            else:
                w.delete(0, tk.END)
        for w in self.emplacement_entries.values():
            w.delete(0, tk.END)
        self.archivage_entry.set('')
        self.file_label.config(text="Aucun fichier sélectionné")
        self.file_path = None
    
    def add_document(self):
        data = {
            'numero_dossier': self.entries['numero_du_dossier'].get().strip(),
            'modele'        : self.entries['modele'].get().strip(),
            'langue'        : self.entries['langue'].get().strip(),
            'titer'         : self.entries['titer'].get().strip(),
            'date'          : self.entries['date'].get().strip(),
            'departement'   : self.entries['departement'].get().strip(),
            'objet'         : self.entries['objet'].get().strip(),
            'numero_piece'  : self.entries['numero_de_piece'].get().strip(),
            'type'          : self.entries['type'].get().strip(),
            'statut_archivage': self.archivage_entry.get().strip(),
            'numero_boite'  : self.emplacement_entries['numero_boite'].get().strip(),
            'salle'         : self.emplacement_entries['salle'].get().strip(),
            'etagere'       : self.emplacement_entries['etagere'].get().strip(),
            'rayon'         : self.emplacement_entries['rayon'].get().strip(),
            'file_path'     : self.file_path
        }
        # validations
        if not data['numero_piece'] or not data['objet']:
            messagebox.showerror("Erreur", "Le N° de pièce et l'objet sont obligatoires.")
            return
        
        try:
            self.db.add_document(**data)
            messagebox.showinfo("Succès", "Document ajouté !")
            self.clear_form()
            self.show_all_documents()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter :\n{e}")
    
    def open_selected_file(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Ouvrir", "Sélectionnez d'abord un document.")
            return
        values = self.tree.item(sel[0], "values")
        path = values[-1]  # Le fichier est dans la dernière colonne
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
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile=default
        )
        if not fp: return
        hdrs = [self.tree.heading(c)["text"] for c in self.tree["columns"]]
        rows = [self.tree.item(i,"values") for i in self.tree.get_children()]
        try:
            with open(fp, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(hdrs)
                w.writerows(rows)
            messagebox.showinfo("Export", f"Exporté dans {fp}")
            if messagebox.askyesno("Ouvrir", "Ouvrir le fichier exporté ?"):
                webbrowser.open(fp)
        except Exception as e:
            messagebox.showerror("Erreur d'export", str(e))
    
    def show_dossier_details(self):
        """Affiche les détails d'un dossier spécifique"""
        num = self.detailed_dossier_entry.get().strip()
        if not num:
            messagebox.showwarning("Recherche détaillée", "Entrez un N° de dossier.")
            return
        
        # Rechercher tous les documents de ce dossier
        criteria = {'numero_dossier': num}
        docs = self.db.search_documents(criteria)
        
        if not docs:
            messagebox.showinfo("Recherche détaillée", f"Aucun document pour le dossier {num}.")
            return
        
        # Créer une fenêtre de détails pour le dossier
        self.show_dossier_window(num, docs)
    
    def search_by_objet(self):
        """Recherche par objet et affiche les résultats dans le tableau principal"""
        objet = self.detailed_objet_entry.get().strip()
        if not objet:
            messagebox.showwarning("Recherche par objet", "Entrez un objet à rechercher.")
            return
        
        # Rechercher par objet
        criteria = {'objet': objet}
        rows = self.db.search_documents(criteria)
        self.populate_tree(rows)
        
        if rows:
            messagebox.showinfo("Recherche", f"{len(rows)} document(s) trouvé(s) contenant '{objet}'.")
        else:
            messagebox.showinfo("Recherche", f"Aucun document trouvé contenant '{objet}'.")
    
    def show_dossier_window(self, numero_dossier, docs):
        """Affiche une fenêtre avec tous les documents d'un dossier"""
        pop = tk.Toplevel(self.root)
        pop.title(f"Détails du dossier {numero_dossier}")
        pop.geometry("800x600+150+150")
        
        # Frame principal
        main_frame = ttk.Frame(pop)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Titre
        title_label = ttk.Label(main_frame, text=f"Dossier N° {numero_dossier}", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Informations générales
        info_frame = ttk.LabelFrame(main_frame, text="Informations générales")
        info_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(info_frame, text=f"Nombre de documents: {len(docs)}").pack(anchor="w", padx=10, pady=2)
        
        # Créer un tableau pour les documents du dossier
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)
        
        cols = ("id", "numero_piece", "date", "type", "objet", "statut_archivage")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
        
        headers = {
            "id": "ID",
            "numero_piece": "N° Pièce", 
            "date": "Date",
            "type": "Type",
            "objet": "Objet",
            "statut_archivage": "Statut"
        }
        
        for col in cols:
            tree.heading(col, text=headers[col])
            if col == "id":
                tree.column(col, width=50)
            elif col == "objet":
                tree.column(col, width=200)
            else:
                tree.column(col, width=120)
        
        # Remplir le tableau avec les documents du dossier
        for doc in docs:
            vals = [
                str(doc[0]) if doc[0] is not None else "",     # id
                str(doc[16]) if doc[16] is not None else "",   # numero_piece
                str(doc[5]) if doc[5] is not None else "",     # date
                str(doc[10]) if doc[10] is not None else "",   # type
                str(doc[8]) if doc[8] is not None else "",     # objet
                str(doc[17]) if doc[17] is not None else ""    # statut_archivage
            ]
            tree.insert("", "end", values=vals)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        
        # Boutons d'action
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        def voir_detail_piece():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Sélection", "Sélectionnez d'abord un document.")
                return
            doc_id = tree.item(sel[0], "values")[0]
            doc = self.db.get_document_by_id(doc_id)
            if doc:
                self.show_individual_document_details(doc)
        
        def exporter_dossier():
            # Exporter les documents de ce dossier
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"dossier_{numero_dossier}_{timestamp}.csv"
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")],
                initialfile=filename
            )
            if filepath:
                try:
                    with open(filepath, "w", newline="", encoding="utf-8") as f:
                        writer = csv.writer(f, delimiter=";")
                        writer.writerow(["ID", "N° Pièce", "Date", "Type", "Objet", "Statut"])
                        for doc in docs:
                            writer.writerow([
                                doc[0], doc[16] or "", doc[5] or "", 
                                doc[10] or "", doc[8] or "", doc[17] or ""
                            ])
                    messagebox.showinfo("Export", f"Dossier exporté vers {filepath}")
                except Exception as e:
                    messagebox.showerror("Erreur", f"Impossible d'exporter: {e}")
        
        ttk.Button(btn_frame, text="Voir détail sélectionné", command=voir_detail_piece).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Exporter ce dossier", command=exporter_dossier).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Fermer", command=pop.destroy).pack(side="left", padx=5)

    def show_document_details(self):
        """Affiche les détails complets d'un document par son numéro de pièce"""
        num = self.detailed_search_entry.get().strip()
        if not num:
            messagebox.showwarning("Recherche détaillée", "Entrez un N° de pièce.")
            return
        doc = self.db.get_document_by_numero_piece(num)
        if not doc:
            messagebox.showinfo("Recherche détaillée", f"Aucun document pour la pièce {num}.")
            return
        
        self.show_individual_document_details(doc)
    
    def show_individual_document_details(self, doc):
        """Affiche une fenêtre détaillée pour un document spécifique"""
        pop = tk.Toplevel(self.root)
        pop.title(f"Détails complets - Pièce {doc.get('numero_piece', 'N/A')}")
        pop.geometry("700x750+200+50")
        
        # Frame principal avec scrollbar
        canvas = tk.Canvas(pop)
        scrollbar = ttk.Scrollbar(pop, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Titre principal
        title_frame = ttk.Frame(scrollable_frame)
        title_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(title_frame, text="DÉTAILS COMPLETS DU DOCUMENT", 
                 font=("Arial", 16, "bold")).pack()
        ttk.Separator(title_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # Section Identification
        id_frame = ttk.LabelFrame(scrollable_frame, text="🔍 IDENTIFICATION")
        id_frame.pack(fill="x", padx=20, pady=5)
        
        id_fields = [
            ("ID Système", doc.get("id")),
            ("Numéro de Dossier", doc.get("numero_dossier")),
            ("Numéro de Pièce", doc.get("numero_piece")),
        ]
        
        for i, (label, value) in enumerate(id_fields):
            ttk.Label(id_frame, text=f"{label}:", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky="e", padx=10, pady=3)
            ttk.Label(id_frame, text=value or "Non renseigné", 
                     background="white", relief="sunken").grid(
                row=i, column=1, sticky="ew", padx=10, pady=3)
            id_frame.columnconfigure(1, weight=1)
        
        # Section Contenu
        content_frame = ttk.LabelFrame(scrollable_frame, text="📄 CONTENU")
        content_frame.pack(fill="x", padx=20, pady=5)
        
        content_fields = [
            ("Objet", doc.get("objet")),
            ("Titre", doc.get("titer")),
            ("Type", doc.get("type")),
            ("Modèle", doc.get("modele")),
            ("Langue", doc.get("langue")),
            ("Date", doc.get("date")),
        ]
        
        for i, (label, value) in enumerate(content_fields):
            ttk.Label(content_frame, text=f"{label}:", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky="e", padx=10, pady=3)
            ttk.Label(content_frame, text=value or "Non renseigné", 
                     background="white", relief="sunken", wraplength=400).grid(
                row=i, column=1, sticky="ew", padx=10, pady=3)
            content_frame.columnconfigure(1, weight=1)
        
        # Section Administrative
        admin_frame = ttk.LabelFrame(scrollable_frame, text="🏢 INFORMATIONS ADMINISTRATIVES")
        admin_frame.pack(fill="x", padx=20, pady=5)
        
        admin_fields = [
            ("Département", doc.get("departement")),
            ("Statut d'Archivage", doc.get("statut_archivage")),
        ]
        
        for i, (label, value) in enumerate(admin_fields):
            ttk.Label(admin_frame, text=f"{label}:", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky="e", padx=10, pady=3)
            ttk.Label(admin_frame, text=value or "Non renseigné", 
                     background="white", relief="sunken").grid(
                row=i, column=1, sticky="ew", padx=10, pady=3)
            admin_frame.columnconfigure(1, weight=1)
        
        # Section Emplacement Physique
        location_frame = ttk.LabelFrame(scrollable_frame, text="📍 EMPLACEMENT PHYSIQUE")
        location_frame.pack(fill="x", padx=20, pady=5)
        
        location_fields = [
            ("Numéro de Boîte", doc.get("numero_boite")),
            ("Salle", doc.get("salle")),
            ("Étagère", doc.get("etagere")),
            ("Rayon", doc.get("rayon")),
        ]
        
        for i, (label, value) in enumerate(location_fields):
            ttk.Label(location_frame, text=f"{label}:", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky="e", padx=10, pady=3)
            ttk.Label(location_frame, text=value or "Non renseigné", 
                     background="white", relief="sunken").grid(
                row=i, column=1, sticky="ew", padx=10, pady=3)
            location_frame.columnconfigure(1, weight=1)
        
        # Section Fichier
        file_frame = ttk.LabelFrame(scrollable_frame, text="📁 FICHIER NUMÉRIQUE")
        file_frame.pack(fill="x", padx=20, pady=5)
        
        file_path = doc.get("file_path")
        ttk.Label(file_frame, text="Chemin:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky="e", padx=10, pady=3)
        
        if file_path:
            ttk.Label(file_frame, text=file_path, 
                     background="lightgreen", relief="sunken", wraplength=400).grid(
                row=0, column=1, sticky="ew", padx=10, pady=3)
            
            file_status = "✅ Accessible" if os.path.exists(file_path) else "❌ Introuvable"
            ttk.Label(file_frame, text="Statut:", font=("Arial", 10, "bold")).grid(
                row=1, column=0, sticky="e", padx=10, pady=3)
            ttk.Label(file_frame, text=file_status, 
                     background="lightgreen" if os.path.exists(file_path) else "lightcoral", 
                     relief="sunken").grid(row=1, column=1, sticky="ew", padx=10, pady=3)
        else:
            ttk.Label(file_frame, text="Aucun fichier associé", 
                     background="lightgray", relief="sunken").grid(
                row=0, column=1, sticky="ew", padx=10, pady=3)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Boutons d'action
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.pack(pady=20)
        
        if file_path and os.path.exists(file_path):
            def _open_file(): 
                webbrowser.open(file_path)
            ttk.Button(btn_frame, text="🔗 Ouvrir le fichier", command=_open_file).pack(side="left", padx=5)
        
        def _edit_document():
            pop.destroy()
            self.current_document_id = doc.get("id")
            self.fill_form_with_document(doc)
            self.add_btn.config(state="disabled")
            self.modify_btn.config(state="normal")
            self.cancel_btn.config(state="normal")
            messagebox.showinfo("Modification", "Document chargé pour modification.")
        
        ttk.Button(btn_frame, text="✏️ Modifier", command=_edit_document).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ Fermer", command=pop.destroy).pack(side="left", padx=5)
        
        # Configuration finale
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Permettre le scroll avec la molette
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        pop.bind_all("<MouseWheel>", _on_mousewheel)

if __name__ == "__main__":
    root = tk.Tk()
    app  = GEDApp(root)
    root.mainloop()