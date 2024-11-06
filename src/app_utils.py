import tkinter as tk
from tkinter import ttk, messagebox

def configure_styles(style):
    """Configure les styles pour les boutons et les barres de progression."""
    style.configure("TButton", font=("Helvetica", 12), padding=10, relief="flat", background="#007BFF", foreground="white")
    style.map("TButton", background=[("active", "#0056b3")])
    style.configure("TProgressbar", thickness=15, troughcolor="#D3D3D3", background="#007BFF")

def create_file_selection_section(root, label_text, row, select_command):
    """Crée une section de sélection de fichier."""
    tk.Label(root, text=label_text, font=("Helvetica", 12, "bold"), bg="#f5f5f5").grid(row=row, column=0, padx=10, pady=5, sticky="w")
    entry = tk.Entry(root, width=50)
    entry.grid(row=row, column=1, padx=10, pady=5)
    tk.Button(root, text="Parcourir", command=lambda: entry.insert(0, select_command()), bg="blue", fg="white", relief="raised", font=("Helvetica", 12)).grid(row=row, column=2, padx=10, pady=5, ipadx=10)
    return entry

def create_progress_bars(root):
    """Crée les barres de progression."""
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=500, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=4, padx=25, pady=10, sticky="ew")

    secondary_progress_var = tk.IntVar()
    secondary_progress_bar = ttk.Progressbar(root, variable=secondary_progress_var, maximum=100, style="TProgressbar")
    secondary_progress_bar.grid(row=7, column=0, columnspan=4, padx=25, pady=10, sticky="ew")

    return progress_var, progress_bar, secondary_progress_var, secondary_progress_bar

def validate_entries(doc_entry, excel_entry, img_entry):
    """Vérifie si les entrées obligatoires sont remplies."""
    if not doc_entry.get():
        messagebox.showwarning("Avertissement", "Veuillez sélectionner un document Word.")
        return False
    if not excel_entry.get():
        messagebox.showwarning("Avertissement", "Veuillez sélectionner un fichier Excel.")
        return False
    if not img_entry.get():
        messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier des images.")
        return False
    return True

def convert_dimensions(width_entry, height_entry):
    """Convertit les dimensions des cartes de pixels en pouces."""
    width_pixels = int(width_entry.get()) if width_entry.get().isdigit() else 5.75 * 96
    height_pixels = int(height_entry.get()) if height_entry.get().isdigit() else None
    width_inches = width_pixels / 96
    height_inches = height_pixels / 96 if height_pixels else None
    return width_inches, height_inches

def confirm_parameters():
    """Affiche une boîte de confirmation avec les paramètres sélectionnés."""
    return messagebox.askyesno(
        "Confirmation", 
        f"Confirmez-vous les paramètres actuels ?\n"
    )