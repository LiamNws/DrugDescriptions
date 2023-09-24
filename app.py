import pandas as pd
import os
import pyperclip
import re

excel = os.path.join("Source", "DrugDescriptions.xlsx")

try:
    df = pd.read_excel(excel)
except FileNotFoundError:
    print("File not found.")
    exit(1)
except Exception as e:
    print("An error has occurred while reading the file")
    exit(1)


while True:

    user_input = input("Enter Drug Name (e.g metformin 500): ")

    matching = df[(df["Drug Name"].str.contains(user_input, case=False, na=False)) & (df["Active Status"] != "FALSE")]

    if not matching.empty:
        for i, row in matching.iterrows():

            form = row["Form"]

            if pd.isna(row["Shape"]) or row["Shape"] == "":
                shape = ""
            else:
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
            sentence1 = re.sub(r"(?<!\n)\s+", " ", sentence1)
            sentence2 = f"{name}: {colour} {shape} {form}, no imprint. \n -LN"
            sentence2 = re.sub(r"(?<!\n)\s+", " ", sentence2)

            if imp_found:
                print(sentence1)
                # print("Copied to clipboard.")
                pyperclip.copy(sentence1)
            elif not imp_found:
                print(sentence2)
                # print("Copied to clipboard.")
                pyperclip.copy(sentence2)

    else:
        print(f"No active drug was found for {user_input}")

    cont_loop = input("Ya thirsty for more? (Y/N): ")
    if cont_loop == "N" or cont_loop == "n":
        break
