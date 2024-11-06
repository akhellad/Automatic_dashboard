import os
import time
import win32com.client as win32
from tkinter import messagebox
from PIL import ImageGrab
import excel2img

def autofit_excel_range(sheet, table_range):
    """Applique un Autofit sur les colonnes et les lignes d'une plage de cellules dans une feuille."""
    # Sélectionne la plage spécifiée
    try:
        excel_range = sheet.Range(table_range)
        excel_range.Columns.AutoFit()
        excel_range.Rows.AutoFit()
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'application de l'Autofit : {e}")

def save_excel_graph_or_range(sheet, output_dir, chart_name=None, table_range=None, custom_filename=None, input_excel_path=None):
    """Enregistre un graphique ou une plage de cellules Excel sous forme d'image PNG."""
    if chart_name:
        chart = sheet.Shapes(chart_name)
        chart.Copy()
        image = ImageGrab.grabclipboard()
        if image:
            filename = custom_filename if custom_filename else chart_name
            image_path = os.path.join(output_dir, f"{filename}.png")
            image.save(image_path, 'PNG')

    if table_range:
        autofit_excel_range(sheet, table_range)
        filename = custom_filename if custom_filename else f"plage_{table_range.replace(':', '_')}"
        image_path = os.path.join(output_dir, f"{filename}.png")
        excel2img.export_img(input_excel_path, image_path, sheet.Name, table_range)

def batch_export_excel_elements(input_excel_path, output_dir, elements_to_export, progress_var, step_text, secondary_progress_var, root):
    """Exporte les graphiques et les plages de cellules Excel en tant qu'images PNG."""
    step_text.set("Export des graphiques/tableaux Excel...")
    excel = win32.Dispatch("Excel.Application")
    excel.ScreenUpdating = False
    excel.Visible = False
    excel.DisplayAlerts = False

    # Ouvrir le fichier Excel
    try:
        workbook = excel.Workbooks.Open(input_excel_path)
        total_elements = len(elements_to_export)
        # Parcours des éléments à exporter
        for i, element in enumerate(elements_to_export):
            sheet_name = element.get('sheet_name')
            chart_name = element.get('chart_name')
            table_range = element.get('table_range')
            custom_filename = element.get('custom_filename')
            try:
                sheet = workbook.Sheets(sheet_name)
                save_excel_graph_or_range(sheet, output_dir, chart_name, table_range, custom_filename, input_excel_path)
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement de l'élément '{sheet_name}' : {e}")
            secondary_progress_var.set((i + 1) * 100 / total_elements)
            time.sleep(0.05)
            root.update_idletasks()
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'ouverture du fichier Excel : {e}")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()
        progress_var.set(70)