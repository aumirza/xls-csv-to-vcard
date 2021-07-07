import os.path
import pandas as pd
from string import Template

filename = input("Enter file name : ").strip()

print("Have a name column \n 0. No \n 1. Yes ")
name_col_prompt = int(input("=> ").strip())
if name_col_prompt:
    print("Enter column name or number (starts with 0) for contact name")
    name_col_input = input("=> ").strip()
    name_col = int(name_col_input) or name_col_input
else:
    name_col = ""
    
print("Enter column name or number (starts with 0) for phone number")
num_col_input = input("=> ").strip()
number_col = int(num_col_input) if num_col_input.isnumeric() else num_col_input

extension = os.path.splitext(filename)[1][1:].strip().lower()

if any([ isinstance(name_col, int), isinstance(number_col, int) ]):
    header=None
else:
    header = 0

if extension == "csv":
    df = pd.read_csv(filename)
elif extension == "xlsx":
    df = pd.read_excel(filename,engine='openpyxl',header=header)
else:
    print("Extension is not supported")
    exit()

v_card_template = Template("""
BEGIN:VCARD
VERSION:4.0
FN:$full_name
TEL;TYPE#home,voice;VALUE#uri:tel:$number
END:VCARD
""")

v_card_file = "contacts.vcf"


with open(v_card_file,"w") as f:
    for i in df.index:
        if name_col:
            name = df[name_col][i]
        else:
            name = f"AA-{i}"
        number = df[number_col][i]
        person = v_card_template.substitute(full_name=name, number = number)

        print(person,file=f)

