import tkinter as tk
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
from xlutils.save import save




def enregistrer_panne():
    # Récupérer les données saisies par l'utilisateur
    Nom_client = entry_nom_panne.get()
    description = entry_description.get()
    action = entry_action.get()
    num = entry_num_contact.get()
    
    # date du jour dans date
    date = entry_date.get()

   # Ouvrir le fichier Excel existant ou en créer un nouveau
    try:
        workbook = open_workbook('rapport_pannes.xls', formatting_info=True)
        original_sheet = workbook.sheet_by_index(0)
        row_count = original_sheet.nrows
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Rapport de pannes')
        for row_index in range(row_count):
            for col_index in range(original_sheet.ncols):
                sheet.write(row_index, col_index, original_sheet.cell(row_index, col_index).value)
    except FileNotFoundError:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Rapport de pannes')
        sheet.write(0, 0, 'Numéro de panne')
        sheet.write(0, 1, 'Description de la panne')
        sheet.write(0, 2, 'Date de la panne')
        row_count = 0

    # Ajouter la nouvelle panne au fichier Excel
    sheet.write(row_count, 0, Nom_client)
    sheet.write(row_count, 1, description)
    sheet.write(row_count, 2, date)
    sheet.write(row_count, 3, action)
    sheet.write(row_count, 4, num)
    row_count = 0

    # Enregistrer le fichier Excel
    workbook.save('rapport_pannes.xls')

    # Afficher un message de confirmation
    lbl_confirmation.config(text="Panne enregistrée avec succès dans le fichier Excel.")

# Créer la fenêtre principale de l'application
window = tk.Tk()
window.title("Enregistrement des pannes")
window.geometry("1000x600")


# Créer les étiquettes et les champs de saisie pour les informations de la panne
lbl_nom_panne = tk.Label(window, text="Nom du client :")
lbl_nom_panne.pack()
entry_nom_panne = tk.Entry(window)
entry_nom_panne.pack()

lbl_description = tk.Label(window, text="Description de la panne :")
lbl_description.pack()
entry_description = tk.Entry(window)
entry_description.config(width=50)
entry_description.pack()

lbl_action = tk.Label(window, text="Action effectuée :")
lbl_action.pack()
entry_action = tk.Entry(window)
entry_action.config(width=50)
entry_action.pack()

lbl_date = tk.Label(window, text="Date de la panne :")
lbl_date.pack()
entry_date = tk.Entry(window)
entry_date.pack()

lbl_num_contact = tk.Label(window, text="Numéro de contact :")
lbl_num_contact.pack()
entry_num_contact = tk.Entry(window)
entry_num_contact.pack()

# Créer le bouton pour enregistrer la panne
btn_enregistrer = tk.Button(window, text="Enregistrer", command=enregistrer_panne)
btn_enregistrer.pack()

# Créer une étiquette pour afficher le message de confirmation
lbl_confirmation = tk.Label(window, text="")
lbl_confirmation.pack()

# Créer un tablau qui affiche les pannes enregistrées
# Créer une étiquette pour le titre du tableau
lbl_titre_tableau = tk.Label(window, text="Liste des pannes enregistrées")
lbl_titre_tableau.config(font=("Arial", 14))
lbl_titre_tableau.config(width=200)
lbl_titre_tableau.pack()

# Créer un tableau pour afficher les pannes enregistrées
tableau = tk.Listbox(window)
tableau.pack()

# Ouvrir le fichier Excel qui contient les pannes enregistrées
try:
    workbook = open_workbook('rapport_pannes.xls')
    sheet = workbook.sheet_by_index(0)
    row_count = sheet.nrows
    for row_index in range(row_count):
        tableau.insert(row_index, sheet.cell(row_index, 0).value)
except FileNotFoundError:
    pass

# remplir les étiquttes avec les données de la panne sélectionnée
def afficher_panne(event):
    # Récupérer le numéro de la panne sélectionnée
    selection = tableau.curselection()
    panne_index = selection[0]
    # Ouvrir le fichier Excel qui contient les pannes enregistrées
    workbook = open_workbook('rapport_pannes.xls')
    sheet = workbook.sheet_by_index(0)
    # Afficher les données de la panne sélectionnée
    entry_nom_panne.delete(0, tk.END)
    entry_nom_panne.insert(0, sheet.cell(panne_index, 0).value)
    entry_description.delete(0, tk.END)
    entry_description.insert(0, sheet.cell(panne_index, 1).value)
    entry_date.delete(0, tk.END)
    entry_date.insert(0, sheet.cell(panne_index, 2).value)
    entry_action.delete(0, tk.END)
    entry_action.insert(0, sheet.cell(panne_index, 3).value)
    entry_num_contact.delete(0, tk.END)
    entry_num_contact.insert(0, sheet.cell(panne_index, 4).value)

    # si une information est modifiée, enregistrer la panne modifiée que si une entrée est modifiée
    def enregistrer_panne_modifiee():
        # Récupérer les données de la panne modifiée
        Nom_client = entry_nom_panne.get()
        description = entry_description.get()
        date = entry_date.get()
        action = entry_action.get()
        num = entry_num_contact.get()
        # Ouvrir le fichier Excel qui contient les pannes enregistrées
        workbook = open_workbook('rapport_pannes.xls', formatting_info=True)
        original_sheet = workbook.sheet_by_index(0)
        row_count = original_sheet.nrows
        workbook = xlwt.Workbook()
        # Enregistrer la panne modifiée dans le fichier Excel sans utiliser write
        sheet.cell(panne_index, 0).value = Nom_client
        sheet.cell(panne_index, 1).value = description
        sheet.cell(panne_index, 2).value = date
        sheet.cell(panne_index, 3).value = action
        sheet.cell(panne_index, 4).value = num

        # Enregistrer le fichier Excel
        workbook.save('rapport_pannes.xls')
        # Afficher un message de confirmation
        lbl_confirmation.config(text="Panne enregistrée avec succès dans le fichier Excel.")

    # Créer le bouton pour enregistrer la panne modifiée
    btn_enregistrer_modification = tk.Button(window, text="Enregistrer la modification", command=enregistrer_panne_modifiee)
    btn_enregistrer_modification.pack()



    

# Lier la fonction afficher_panne à l'événement de sélection d'une panne
tableau.bind('<<ListboxSelect>>', afficher_panne)


# Lancer l'application
window.mainloop()
