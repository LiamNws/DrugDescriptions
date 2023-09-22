import pandas as pd
import openpyxl
import os
import pyperclip

excel = os.path.join("Source", "DrugDescriptions.xlsx")

try:
    df = pd.read_excel(excel)
except FileNotFoundError:
    print("File not found")
    exit(1)
except Exception as e:
    print("An error has occurred while reading the file")
    exit(1)

startstop = True
while startstop:

    user_input = input("Enter Drug Name: ")

    matching = df[df["Drug Name"].str.contains(user_input, case=False, na=False)]

    if not matching.empty:
        for i, row in matching.iterrows():
            form = row["Form"]
            shape = row["Shape"]
            colour = row["Colour"]

            if row["Imprint"] == "none":
                imprint = ""
                imp_found = False
            else:
                imprint = row["Imprint"]
                imp_found = True

            name = row["Drug Name"]

            sentence1 = f"{name}: {colour} {shape} {form}, '{imprint}' imprinted. \n -LN"
            sentence2 = f"{name}: {colour} {shape} {form}, no imprint. \n -LN"

            if imp_found:
                print(sentence1)
                pyperclip.copy(sentence1)
            elif not imp_found:
                print(sentence2)
                pyperclip.copy(sentence2)



    else:
        print(f"{user_input} was not found.")

    cont_loop = input("Ya thirsty for more? (Y/N): ")
    if cont_loop == "N" or cont_loop == "n":
        startstop = False
