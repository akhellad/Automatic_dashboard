import os
import json
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, ttk
from docx import Document
from config import JSON_PATHS, elements_to_export, image_data
from utils import trouver_orphelins_initiaux, normalize_bookmark_name, ajouter_nouvelle_carte, afficher_cartes

def open_combined_options_window(doc_path, images_folder, root):
    """Ouvre une fenêtre combinée avec deux onglets pour les options de gestion des cartes et des graphiques/tableaux."""

    # Vérifie si les entrées requises sont fournies
    if not doc_path:
        messagebox.showwarning("Avertissement", "Veuillez sélectionner un document Word.")
        return
    if not images_folder:
        messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'images.")
        return

    # Créer la fenêtre combinée d'options
    combined_window = tk.Toplevel(root)
    combined_window.title("Options - Gérer les éléments")
    combined_window.geometry("890x700")
    combined_window.config(bg="#f5f5f5")

    # Création du Notebook pour les onglets
    notebook = ttk.Notebook(combined_window)
    notebook.pack(expand=True, fill="both")

    # Onglet pour les cartes avec barre de défilement
    maps_frame = tk.Frame(notebook, bg="#f5f5f5")
    notebook.add(maps_frame, text="Cartes")
    
    maps_canvas = tk.Canvas(maps_frame, bg="#f5f5f5")
    maps_scrollbar = tk.Scrollbar(maps_frame, orient="vertical", command=maps_canvas.yview)
    maps_scrollable_frame = tk.Frame(maps_canvas, bg="#f5f5f5")

    maps_scrollable_frame.bind(
        "<Configure>",
        lambda e: maps_canvas.configure(scrollregion=maps_canvas.bbox("all"))
    )

    maps_canvas.create_window((0, 0), window=maps_scrollable_frame, anchor="nw")
    maps_canvas.configure(yscrollcommand=maps_scrollbar.set)

    maps_canvas.pack(side="left", fill="both", expand=True)
    maps_scrollbar.pack(side="right", fill="y")

    # Titre et bouton pour ajouter une carte
    title_frame_maps = tk.Frame(maps_scrollable_frame, bg="#f5f5f5")
    title_frame_maps.pack(fill="x", pady=10)
    tk.Label(title_frame_maps, text="Cartes", font=("Helvetica", 12, "bold"), bg="#f5f5f5").pack(side="left")
    add_button_maps = tk.Button(title_frame_maps, text="Ajouter une carte", command=lambda: ajouter_nouvelle_carte(doc_path, images_folder, combined_window), bg="blue", fg="white", relief="raised", font=("Helvetica", 12))
    add_button_maps.pack(side="right", padx=5)

    # Cadre pour la liste des cartes
    list_frame_maps = tk.Frame(maps_scrollable_frame, bg="#f5f5f5")
    list_frame_maps.pack(fill="both", expand=True, padx=10, pady=10)
    afficher_cartes(list_frame_maps, doc_path)

    # Onglet pour les graphiques/tableaux avec barre de défilement
    graphs_frame = tk.Frame(notebook, bg="#f5f5f5")
    notebook.add(graphs_frame, text="Graphiques/Tableaux")
    
    graphs_canvas = tk.Canvas(graphs_frame, bg="#f5f5f5")
    graphs_scrollbar = tk.Scrollbar(graphs_frame, orient="vertical", command=graphs_canvas.yview)
    graphs_scrollable_frame = tk.Frame(graphs_canvas, bg="#f5f5f5")

    graphs_scrollable_frame.bind(
        "<Configure>",
        lambda e: graphs_canvas.configure(scrollregion=graphs_canvas.bbox("all"))
    )

    graphs_canvas.create_window((0, 0), window=graphs_scrollable_frame, anchor="nw")
    graphs_canvas.configure(yscrollcommand=graphs_scrollbar.set)

    graphs_canvas.pack(side="left", fill="both", expand=True)
    graphs_scrollbar.pack(side="right", fill="y")

    # Titre et bouton pour ajouter un graphique/tableau
    title_frame_graphs = tk.Frame(graphs_scrollable_frame, bg="#f5f5f5")
    title_frame_graphs.pack(fill="x", pady=5)
    tk.Label(title_frame_graphs, text="Graphiques/Tableaux", font=("Helvetica", 12, "bold"), bg="#f5f5f5").pack(side="left", padx=10)
    add_button_graphs = tk.Button(title_frame_graphs, text="Ajouter un graphique/tableau", command=lambda: ajouter_element(doc_path, root),bg="blue", fg="white", relief="raised", font=("Helvetica", 12))
    add_button_graphs.pack(side="right", padx=10)

    # Affichage des graphiques/tableaux
    global options_window, scrollable_frame, orphelins_existants
    orphelins_existants = trouver_orphelins_initiaux(doc_path)
    scrollable_frame = graphs_scrollable_frame  # Remplace par la nouvelle frame de défilement

    # Ajouter chaque élément graphique ou tableau en vérifiant s'il doit être en rouge
    for i, item in enumerate(elements_to_export):
        bookmark_name_normalized = normalize_bookmark_name(image_data[i]['bookmark_name']) if i < len(image_data) else ""
        rouge = bookmark_name_normalized in orphelins_existants
        ajouter_element_interface(i, item, root, rouge=rouge)

def ajouter_element(doc_path, root):
    """Ouvre une fenêtre pour ajouter un nouvel élément graphique ou tableau avec sélection de l'emplacement."""
    modif_window = tk.Toplevel(root)
    modif_window.title("Ajouter un élément")
    modif_window.geometry("550x300")
    modif_window.config(bg="#f5f5f5")

    def verifier_cible(target, doc_path):
        """Vérifie si le `target` existe dans le document Word ou en tant que signet."""
        doc_path = os.path.normpath(doc_path)
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        try:
            document = word_app.Documents.Open(doc_path)
            for paragraph in document.Paragraphs:
                if target.lower() == paragraph.Range.Text.strip().lower():
                    return True
            for bookmark in document.Bookmarks:
                if normalize_bookmark_name(bookmark.Name) == normalize_bookmark_name(target):
                    return True
            return False
        finally:
            document.Close(SaveChanges=False)
            word_app.Quit()
    
    def display_target_selection(doc_path, on_selection, root):
        """Affiche une liste de paragraphes et signets dans le document Word pour sélectionner un emplacement de cible."""
        selection_window = tk.Toplevel(root)
        selection_window.title("Sélectionner l'emplacement de la cible")
        selection_window.geometry("500x400")
        selection_window.config(bg="#f5f5f5")

        # Cadre pour faire défiler la liste des paragraphes et signets
        list_frame = tk.Frame(selection_window, bg="#f5f5f5")
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Ajouter la barre de défilement
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        # Cadre pour les éléments avec barre de défilement
        canvas = tk.Canvas(list_frame, bg="#f5f5f5", yscrollcommand=scrollbar.set)
        scrollbar.config(command=canvas.yview)
        canvas.pack(side="left", fill="both", expand=True)

        # Ajouter un cadre pour afficher les paragraphes et signets dans le canvas
        inner_frame = tk.Frame(canvas, bg="#f5f5f5")
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Charger le document Word
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        document = word_app.Documents.Open(os.path.normpath(doc_path))

        try:
            # Afficher les paragraphes
            tk.Label(inner_frame, text="Paragraphes :", bg="#f5f5f5", font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=(5, 2))
            for i, paragraph in enumerate(document.Paragraphs):
                if paragraph.Range.Text.strip():
                    text = paragraph.Range.Text.strip()
                    frame = tk.Frame(inner_frame, bg="#f5f5f5")
                    frame.pack(fill="x", pady=2)
                    tk.Label(frame, text=f"Paragraphe {i+1}: {text[:50]}...", bg="#f5f5f5").pack(side="left")
                    tk.Button(frame, text="Sélectionner", command=lambda text=text: on_selection(text)).pack(side="right")

            # Afficher les signets
            tk.Label(inner_frame, text="Signets :", bg="#f5f5f5", font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
            for bookmark in document.Bookmarks:
                frame = tk.Frame(inner_frame, bg="#f5f5f5")
                frame.pack(fill="x", pady=2)
                tk.Label(frame, text=f"Signet: {bookmark.Name}", bg="#f5f5f5").pack(side="left")
                tk.Button(frame, text="Sélectionner", command=lambda name=bookmark.Name: on_selection(name)).pack(side="right")

            inner_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))

        finally:
            document.Close(SaveChanges=False)
            word_app.Quit()
        
        selection_window.mainloop()

    # Fonction pour afficher la liste des paragraphes et signets et sélectionner un emplacement
    def select_target(root):
        def on_selection(selected_target):
            # Supprime le texte existant dans le champ cible d'insertion
            target_entry.delete(0, tk.END)
            # Insère le texte sélectionné dans le champ cible d'insertion
            target_entry.insert(0, selected_target)
            # Ferme la fenêtre de sélection après la sélection

        # Appelle la fonction de sélection avec la fonction de callback `on_selection`
        display_target_selection(doc_path, on_selection, root)

    def save_new_element():
        sheet_name = sheet_name_entry.get()
        table_range = table_range_entry.get()
        chart_name = chart_name_entry.get()
        custom_filename = custom_filename_entry.get()
        target = target_entry.get()
        image_path = image_path_entry.get()
        bookmark_name = bookmark_name_entry.get()

        if not sheet_name or not custom_filename or not target or not bookmark_name:
            messagebox.showwarning("Avertissement", "Veuillez remplir tous les champs obligatoires.")
            return
        if any(el.get("custom_filename") == custom_filename for el in elements_to_export):
            messagebox.showwarning("Avertissement", "Un élément avec ce nom existe déjà.")
            return
        if not verifier_cible(target, doc_path):
            messagebox.showwarning("Avertissement", "La cible n'existe ni dans le document Word, ni parmi les signets.")
            return
        if bool(table_range) == bool(chart_name):
            messagebox.showwarning("Avertissement", "Veuillez remplir uniquement 'Plage de table' ou 'Nom du graphique', pas les deux.")
            return

        new_element = {"sheet_name": sheet_name, "custom_filename": custom_filename}
        if table_range:
            new_element["table_range"] = table_range
        elif chart_name:
            new_element["chart_name"] = chart_name
        elements_to_export.append(new_element)

        new_image_data = {
            "target": target,
            "image_path": image_path,
            "bookmark_name": bookmark_name
        }
        image_data.append(new_image_data)

        with open(JSON_PATHS[0], "w", encoding="utf-8") as f:
            json.dump(image_data, f, indent=4, ensure_ascii=False)
        with open(JSON_PATHS[1], "w", encoding="utf-8") as f:
            json.dump(elements_to_export, f, indent=4, ensure_ascii=False)

        messagebox.showinfo("Succès", "L'élément a été ajouté avec succès.")
        rafraichir_fenetre_options()
        modif_window.destroy()

    # Création des champs de saisie
    tk.Label(modif_window, text="Nom de la feuille :", bg="#f5f5f5").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    sheet_name_entry = tk.Entry(modif_window, width=30)
    sheet_name_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Plage de table (ou)", bg="#f5f5f5").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    table_range_entry = tk.Entry(modif_window, width=30)
    table_range_entry.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Nom du graphique :", bg="#f5f5f5").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    chart_name_entry = tk.Entry(modif_window, width=30)
    chart_name_entry.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Nom personnalisé du fichier :", bg="#f5f5f5").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    custom_filename_entry = tk.Entry(modif_window, width=30)
    custom_filename_entry.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Cible d'insertion :", bg="#f5f5f5").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    target_entry = tk.Entry(modif_window, width=20)
    target_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")
    select_target_button = tk.Button(modif_window, text="Sélectionner l'emplacement", command=lambda: select_target(root), bg="#007BFF", fg="white", relief="flat")
    select_target_button.grid(row=4, column=2, padx=5, pady=5)

    tk.Label(modif_window, text="Chemin de l'image :", bg="#f5f5f5").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    image_path_entry = tk.Entry(modif_window, width=30)
    image_path_entry.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Nom du signet :", bg="#f5f5f5").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    bookmark_name_entry = tk.Entry(modif_window, width=30)
    bookmark_name_entry.grid(row=6, column=1, padx=10, pady=5)

    save_button = tk.Button(modif_window, text="Ajouter", command=save_new_element, bg="#007BFF", fg="white", relief="flat")
    save_button.grid(row=7, column=0, columnspan=3, pady=10)

def rafraichir_fenetre_options(root):
    """Rafraîchit le contenu de la fenêtre d'options pour afficher les modifications en temps réel."""
    global scrollable_frame

    # Efface le contenu actuel de la fenêtre
    for widget in scrollable_frame.winfo_children()[1:]:  # Ignorer le titre
        widget.destroy()

    # Ajoute chaque élément graphique ou tableau avec les mises à jour
    for i, item in enumerate(elements_to_export):
        bookmark_name_normalized = normalize_bookmark_name(image_data[i]['bookmark_name']) if i < len(image_data) else ""
        rouge = bookmark_name_normalized in orphelins_existants
        ajouter_element_interface(i, item, root, rouge=rouge)

def ajouter_element_interface(index, item, root, rouge=False):
    """Ajoute un élément graphique ou tableau dans l'interface avec les détails spécifiés."""
    frame = tk.Frame(scrollable_frame, bg="#f5f5f5", padx=10, pady=5, bd=1, relief="solid")
    frame.pack(fill="x", padx=10, pady=5)

    # Utiliser l'élément équivalent de image_data pour le champ "target" et "bookmark_name"
    target_name = image_data[index]['target'] if index < len(image_data) else "Non spécifié"
    bookmark_name_normalized = normalize_bookmark_name(image_data[index]['bookmark_name']) if index < len(image_data) else ""

    # Préparer les labels pour chaque champ
    couleur = "red" if rouge else "#333"
    signet_label = tk.Label(frame, text="Signet : ", font=("Helvetica", 10, "bold"), bg="#f5f5f5", fg=couleur)
    signet_value = tk.Label(frame, text=normalize_bookmark_name(image_data[index]['bookmark_name']), bg="#f5f5f5", fg=couleur)

    if 'table_range' in item:
        type_label = tk.Label(frame, text="Plage : ", font=("Helvetica", 10, "bold"), bg="#f5f5f5", fg=couleur)
        type_value = tk.Label(frame, text=item['table_range'], bg="#f5f5f5", fg=couleur)
    elif 'chart_name' in item:
        type_label = tk.Label(frame, text="Graphique : ", font=("Helvetica", 10, "bold"), bg="#f5f5f5", fg=couleur)
        type_value = tk.Label(frame, text=item['chart_name'], bg="#f5f5f5", fg=couleur)

    location_label = tk.Label(frame, text="Emplacement d'insertion : ", font=("Helvetica", 10, "bold"), bg="#f5f5f5", fg=couleur)
    location_value = tk.Label(frame, text=target_name, bg="#f5f5f5", fg=couleur)

    # Boutons "Supprimer" et "Modifier"
    delete_button = tk.Button(frame, text="Supprimer", command=lambda: supprimer_element(index, frame, root), bg="red", fg="white", relief="raised")
    modify_button = tk.Button(frame, text="Modifier", command=lambda: modifier_element(index, root), bg="blue", fg="white", relief="raised")

    # Placer les éléments dans le frame avec de l'espace entre les paires
    signet_label.grid(row=0, column=0, sticky="w")
    signet_value.grid(row=0, column=1, sticky="w")
    type_label.grid(row=1, column=0, sticky="w")
    type_value.grid(row=1, column=1, sticky="w")
    location_label.grid(row=2, column=0, sticky="w")
    location_value.grid(row=2, column=1, sticky="w")
    delete_button.grid(row=0, column=2, rowspan=3, padx=5)
    modify_button.grid(row=0, column=3, rowspan=3, padx=5)

# Liste globale pour retenir les signets orphelins à conserver en rouge
orphelins_existants = set()

def supprimer_element(index, frame_to_remove, root):
    global orphelins_existants
    
    # Récupère le signet de l'élément supprimé
    deleted_bookmark = normalize_bookmark_name(image_data[index]['bookmark_name'])

    # Supprime l'élément des deux listes
    elements_to_export.pop(index)
    image_data.pop(index)

    # Supprime le widget de l'interface
    frame_to_remove.pack_forget()
    frame_to_remove.destroy()

    # Mettre à jour les dépendances en cascade en ajoutant les nouveaux orphelins
    mettre_a_jour_dependances(deleted_bookmark, root)
    
    # Écrire les modifications dans les fichiers JSON
    with open(JSON_PATHS[0], "w", encoding="utf-8") as f:
        json.dump(image_data, f, indent=4, ensure_ascii=False)
    
    with open(JSON_PATHS[1], "w", encoding="utf-8") as f:
        json.dump(elements_to_export, f, indent=4, ensure_ascii=False)

def mettre_a_jour_dependances(supprime_bookmark, root):
    """Met à jour les dépendances en cascade, marquant en rouge les éléments dépendants de l'élément supprimé."""
    global orphelins_existants

    # Effacer l'interface
    for widget in scrollable_frame.winfo_children()[1:]:  # Garder le titre en premier
        widget.destroy()

    # Ensemble temporaire pour contenir les nouveaux signets à marquer en rouge
    rouge_cascade = set()
    signets_a_verifier = [supprime_bookmark]

    # Vérification en cascade
    while signets_a_verifier:
        current_target = signets_a_verifier.pop()
        rouge_cascade.add(current_target)

        # Rechercher toutes les correspondances de current_target dans les targets
        for item in image_data:
            if 'target' in item and normalize_bookmark_name(item['target']) == current_target:
                dependent_bookmark = normalize_bookmark_name(item['bookmark_name'])
                
                # Ajouter ce signet à vérifier si non déjà marqué
                if dependent_bookmark not in rouge_cascade:
                    signets_a_verifier.append(dependent_bookmark)

    # Ajouter les nouveaux orphelins découverts à la liste globale `orphelins_existants`
    orphelins_existants.update(rouge_cascade)

    # Afficher les éléments, en rouge si leur `bookmark_name` est dans `orphelins_existants`
    for i, item in enumerate(elements_to_export):
        bookmark_name_normalized = normalize_bookmark_name(image_data[i]['bookmark_name']) if i < len(image_data) else ""
        rouge = bookmark_name_normalized in orphelins_existants
        ajouter_element_interface(i, item, root, rouge=rouge)

def modifier_element(index, root):
    """Ouvre une fenêtre pour modifier les caractéristiques de l'élément sélectionné."""
    # Créer une nouvelle fenêtre de modification
    modif_window = tk.Toplevel(root)
    modif_window.title("Modifier l'élément")
    modif_window.geometry("400x450")
    modif_window.config(bg="#f5f5f5")

    # Obtenir les données de l'élément
    element = elements_to_export[index]
    image_data_element = image_data[index]

    # Fonction de sauvegarde des modifications
    def save_changes():
        # Mise à jour des valeurs dans l'élément
        element["sheet_name"] = sheet_name_entry.get()
        
        # Vérification du type d'élément et mise à jour en conséquence
        if "table_range" in element:
            element["table_range"] = table_range_entry.get() or element["table_range"]
        elif "chart_name" in element:
            element["chart_name"] = chart_name_entry.get() or element["chart_name"]
        
        element["custom_filename"] = custom_filename_entry.get()

        image_data_element["target"] = target_entry.get()
        image_data_element["image_path"] = image_path_entry.get()
        image_data_element["bookmark_name"] = bookmark_name_entry.get()

        # Sauvegarder les modifications dans les fichiers JSON
        with open(JSON_PATHS[0], "w", encoding="utf-8") as f:
            json.dump(image_data, f, indent=4, ensure_ascii=False)
        
        with open(JSON_PATHS[1], "w", encoding="utf-8") as f:
            json.dump(elements_to_export, f, indent=4, ensure_ascii=False)

        messagebox.showinfo("Sauvegarde", "Les modifications ont été sauvegardées avec succès.")
        rafraichir_fenetre_options(root)  # Appeler la fonction pour rafraîchir l'interface
        modif_window.destroy()  # Fermer la fenêtre de modification

    # Création des champs modifiables pour chaque propriété
    tk.Label(modif_window, text="Nom de la feuille :", bg="#f5f5f5").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    sheet_name_entry = tk.Entry(modif_window, width=30)
    sheet_name_entry.insert(0, element.get("sheet_name", ""))
    sheet_name_entry.grid(row=0, column=1, padx=10, pady=5)

    if "table_range" in element:
        tk.Label(modif_window, text="Plage de table :", bg="#f5f5f5").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        table_range_entry = tk.Entry(modif_window, width=30)
        table_range_entry.insert(0, element.get("table_range", ""))
        table_range_entry.grid(row=1, column=1, padx=10, pady=5)
    elif "chart_name" in element:
        tk.Label(modif_window, text="Nom du graphique :", bg="#f5f5f5").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        chart_name_entry = tk.Entry(modif_window, width=30)
        chart_name_entry.insert(0, element.get("chart_name", ""))
        chart_name_entry.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Nom personnalisé du fichier :", bg="#f5f5f5").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    custom_filename_entry = tk.Entry(modif_window, width=30)
    custom_filename_entry.insert(0, element.get("custom_filename", ""))
    custom_filename_entry.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Cible d'insertion :", bg="#f5f5f5").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    target_entry = tk.Entry(modif_window, width=30)
    target_entry.insert(0, image_data_element.get("target", ""))
    target_entry.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Chemin de l'image :", bg="#f5f5f5").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    image_path_entry = tk.Entry(modif_window, width=30)
    image_path_entry.insert(0, image_data_element.get("image_path", ""))
    image_path_entry.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(modif_window, text="Nom du signet :", bg="#f5f5f5").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    bookmark_name_entry = tk.Entry(modif_window, width=30)
    bookmark_name_entry.insert(0, image_data_element.get("bookmark_name", ""))
    bookmark_name_entry.grid(row=5, column=1, padx=10, pady=5)

    save_button = tk.Button(modif_window, text="Sauvegarder", command=save_changes, bg="blue", fg="white", relief="raised", font=("Helvetica", 12))
    save_button.grid(row=6, column=0, columnspan=2, pady=10)