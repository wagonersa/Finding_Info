import os
import tkinter
from tkinter import *
from tkinter import filedialog
import openpyxl as op
import pandas as pd
from tabulate import tabulate
master = Tk()

while True:
    print("Please select csv file")
    in_path = tkinter.filedialog.askopenfilenames()
    in_path = ''.join(in_path)
    in_path = re.sub('[(),]', '', in_path)
    df = pd.read_csv(in_path)

    # print("Please select output location")
    # out_path = tkinter.filedialog.askdirectory()
    # out_path = ''.join(out_path)
    # out_path = re.sub('[(),]', '', out_path)
    # print(out_path)

    book = op.Workbook()
    title = ("test.xlsx")
    print(title)
    book.save(title)
    writer = pd.ExcelWriter(title, engine='openpyxl', mode='a')
    writer.book = book

    header = list(df)

    for x in header:
        Matrix = df.pivot_table(index=["Name"], aggfunc='size')
        Matrix.to_excel(writer, "Name")
    writer.save()

    df = pd.read_excel(title, sheet_name="Name")

    new_df = df["Name"].str.split("_", n=1, expand=True)
    df["Condition"] = new_df[0]
    df["Ob_ID"] = new_df[1]
    df["Number_of"] = df[0]
    df.drop(columns=["Name", 0], inplace=True)
    # Condition Numbers
    conditions = [
        '1a', '1b', '1c', '2a', '2b', '2c', '3a', '3b', '3c', '3d', '', '4', '', '5', '6', '7', '8', '9', '10', '11',
        '12',
        '13', '', ''
    ]

    #  What we are replacing with.
    headers = [r'Crack <.25 inch', r'Crack .25 to <.5 inch', r'Crack =>.5 inch or more',
               r'Vertical Displacement < .5 inch',
               r'Vertical Displacement .5 to < =1 inch', r'Vertical Displacement > 1 inch',
               r'Surface Condition Overgrowth',
               r'Surface Condition Spalling / Scaling', r'Surface Condition Ponding',
               r'Surface Condition Obstruction (specify in comments)', r'Curb repair needed (lin. ft)',
               r'No. of Driveways out of compliance', r'Address of each tree', r'No. of drop inlets',
               r'No. of curb inlets',
               r'No. of manholes', r'No. of streetname tiles', r'No. of flumes',
               r'No. of driveways along sidewalk segment',
               r'No. of Fire hydrants', r'No. of Water Meter and Water Valve Adjust', r'No. of trees',
               r'Brick border (lin. ft)', r'Length of retaining wall'
               ]

    for index, elem in enumerate(conditions):
        df.replace(elem, headers[index], inplace=True)
    # df.to_excel(out_path + "/Info_" + object_ID + ".xlsx")
    print(tabulate(df, headers='keys', tablefmt='psql'))
    os.remove(title)
    while True:
        answer = input('Run again? (y/n): ')
        if answer in ('y', 'n'):
            break
        print('Invalid input.')
    if answer == 'y':
        continue
    else:
        print('The program is complete')
        break
