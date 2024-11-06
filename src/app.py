from tkinter import filedialog, messagebox, ttk
from tkinter.ttk import Style
import tkinter as tk
import threading
import os
import json
import shutil
from excel_utils import batch_export_excel_elements
from word_utils import insert_images, insert_maps_to_doc
from config import BASE_IMAGE_PATH, elements_to_export, image_data
from ui_elements import open_combined_options_window
from app_utils import configure_styles, create_file_selection_section, create_progress_bars, validate_entries, convert_dimensions, confirm_parameters

def create_main_buttons(root, doc_entry, excel_entry, img_entry, width_entry, height_entry, progress_var, step_text, secondary_progress_var):
    """Crée les boutons principaux."""
    def confirm_and_start():
        if not validate_entries(doc_entry, excel_entry, img_entry):
            return
        width_inches, height_inches = convert_dimensions(width_entry, height_entry)
        if confirm_parameters(doc_entry, excel_entry, img_entry, width_inches, height_inches):
            start_thread(
                doc_entry.get(),
                img_entry.get(),
                excel_entry.get(), 
                progress_var,
                step_text,
                secondary_progress_var, 
                root, 
                width_inches, 
                height_inches
            )

    tk.Button(
        root, 
        text="Commencer", 
        command=confirm_and_start, 
        bg="green", fg="white", relief="raised", font=("Helvetica", 16)
    ).grid(row=8, column=0, columnspan=4, padx=25, pady=20, sticky='ew')

    tk.Button(
        root, 
        text="Options", 
        command=lambda: open_combined_options_window(doc_entry.get(), img_entry.get(), root), 
        bg="red", fg="white", relief="raised", font=("Helvetica", 16)
    ).grid(row=9, column=0, columnspan=4, padx=25, pady=20, sticky='ew')

def create_root_window():
    """Crée la fenêtre principale de l'application."""
    root = tk.Tk()
    root.title("Insertion de cartes et données")
    root.geometry("800x500")
    root.config(bg="#f5f5f5")
    style = Style()
    style.theme_use("clam")
    configure_styles(style)
    root.bind("<Configure>", lambda event: adjust_progress_bars(root))
    return root

def create_dimension_input_section(root, label_text, row, default_value):
    """Crée une section pour saisir une dimension (largeur/hauteur)."""
    tk.Label(root, text=label_text, font=("Helvetica", 12, "bold"), bg="#f5f5f5").grid(row=row, column=0, padx=10, pady=5, sticky="w")
    frame = tk.Frame(root, bg="#f5f5f5")
    frame.grid(row=row, column=1, padx=10, pady=5, sticky="w")
    entry = tk.Entry(frame, width=10)
    entry.insert(0, default_value)
    entry.pack(side="left", padx=(0, 5))
    tk.Button(frame, text="-", command=lambda: adjust_value(entry, -10)).pack(side="left", padx=2)
    tk.Button(frame, text="+", command=lambda: adjust_value(entry, 10)).pack(side="left", padx=2)
    return entry

def run_app():
    """Point d'entrée principal de l'application."""
    global progress_bar, secondary_progress_bar
    root = create_root_window()

    doc_entry = create_file_selection_section(root, "Sélectionner le fichier Word:", 0, select_word_file)
    excel_entry = create_file_selection_section(root, "Sélectionner le fichier Excel:", 1, select_excel_file)
    img_entry = create_file_selection_section(root, "Sélectionner le dossier des images:", 2, select_images_folder)

    width_entry = create_dimension_input_section(root, "Largeur des cartes (en pixels) :", 5, str(int(5.75 * 96)))
    height_entry = create_dimension_input_section(root, "Hauteur des cartes (en pixels) :", 6, str(int(5.75 * 96)))

    step_text = tk.StringVar()
    tk.Label(root, textvariable=step_text, font=("Helvetica", 10, "italic"), bg="#f5f5f5", fg="#555").grid(row=7, column=0, columnspan=4, padx=10, pady=10)

    progress_var, progress_bar, secondary_progress_var, secondary_progress_bar = create_progress_bars(root)

    create_main_buttons(root, doc_entry, excel_entry, img_entry, width_entry, height_entry, progress_var, step_text, secondary_progress_var)

    root.mainloop()

def adjust_value(entry, increment):
    """Ajuste la valeur de l'entrée (en pixels) en ajoutant ou soustrayant l'incrément."""
    try:
        current_value = int(entry.get())
    except ValueError:
        current_value = 0  # Définit la valeur par défaut si l'entrée est vide ou invalide
    new_value = max(0, current_value + increment)  # Empêche les valeurs négatives
    entry.delete(0, tk.END)
    entry.insert(0, str(new_value))


def start_processing(doc_path, images_folder, input_excel_path, progress_var, step_text, secondary_progress_var, root, width_entry, height_entry):
    try:
        if not os.path.exists(BASE_IMAGE_PATH):
            os.makedirs(BASE_IMAGE_PATH)

        # Placeholder pour tes fonctions comme insert_maps_to_doc, batch_export_excel_elements, etc.
        insert_maps_to_doc(doc_path, images_folder, progress_var, step_text, secondary_progress_var, root, width_entry, height_entry)
        batch_export_excel_elements(input_excel_path, BASE_IMAGE_PATH, elements_to_export, progress_var, step_text, secondary_progress_var, root)
        insert_images(doc_path, image_data, progress_var, step_text, secondary_progress_var, root)

        shutil.rmtree(BASE_IMAGE_PATH)
        messagebox.showinfo("Succès", "Processus terminé avec succès!")
    except Exception as e:
        messagebox.showerror("Erreur", str(e))


def start_thread(doc_path, images_folder, input_excel_path, progress_var, step_text, secondary_progress_var, root, width_entry, height_entry):
    # Vérification des entrées
    if not doc_path:
        messagebox.showwarning("Erreur", "Veuillez sélectionner un fichier Word.")
        return
    if not images_folder:
        messagebox.showwarning("Erreur", "Veuillez sélectionner le dossier des cartes.")
        return
    if not input_excel_path:
        messagebox.showwarning("Erreur", "Veuillez sélectionner un fichier Excel.")
        return

    # Si toutes les entrées sont valides, lancez le traitement dans un thread
    thread = threading.Thread(target=start_processing, args=(doc_path, images_folder, input_excel_path, progress_var, step_text, secondary_progress_var, root, width_entry, height_entry))
    thread.start()


def select_word_file():
    return filedialog.askopenfilename(title="Sélectionner le fichier TdB", filetypes=[("Fichiers Word", "*.docx")])


def select_excel_file():
    return filedialog.askopenfilename(title="Sélectionner le fichier Excel", filetypes=[("Fichiers Excel", "*.xlsx")])


def select_images_folder():
    return filedialog.askdirectory(title="Sélectionner le dossier des cartes")

def select_json_file():
    """Ouvre une boîte de dialogue pour sélectionner un fichier JSON."""
    return filedialog.askopenfilename(title="Sélectionner un fichier JSON", filetypes=[("Fichiers JSON", "*.json")])


def adjust_progress_bars(root):
    window_width = root.winfo_width() - 50
    progress_bar.config(length=window_width)
    secondary_progress_bar.config(length=window_width)