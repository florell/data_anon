import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import subprocess

file_path = ""
processed_file_path = ""
df = None

def k_anonimity(data, columns):
    counter = {}
    for index, row in data.iterrows():
        key = tuple(row[columns])
        if key in counter:
            counter[key] += 1
        else:
            counter[key] = 1
    
    k_values = sorted(counter.values(), reverse=False)[:5]
    return k_values




def card_masking(data: pd.DataFrame):
    for i in range(len(data.index)):
        temp = data['Карта оплаты'][i]
        data.loc[i, 'Карта оплаты'] = temp[0] + '*'*3 + ' ' + '*'*4 + ' ' + '*'*4 + ' ' + '*'*4

def seat_removal(data: pd.DataFrame):
    for i in range(len(data.index)):
        temp = data.at[i, 'Вагон и место']
        data.at[i, 'Вагон и место'] = temp.split('-')[0]
    data.rename(columns={'Вагон и место': 'Вагон'}, inplace=True)

def race_range(data: pd.DataFrame):
    ranges = ['1-150', '151-298', '301-450', '451-598', '701-750', '751-788',]
    for i in range(len(data.index)):
        temp = data['Рейс'][i]
        for j in range(len(ranges)):
            lower, upper = map(int, ranges[j].split('-'))
            if lower <= temp <= upper:
                data.loc[i, 'Рейс'] = ranges[j]
                break

def price_range(data: pd.DataFrame):
    ranges = ['0-500', '500-1000', '1000-1500', '1500-2000', '2000-2500', '2500-3000', '3000-10000000',]
    for i in range(len(data.index)):
        temp = data['Стоимость'][i]
        for j in range(len(ranges)):
            lower, upper = map(int, ranges[j].split('-'))
            if lower <= temp < upper:
                if upper == 10000000:
                    data.loc[i, 'Стоимость'] = str(lower) + '+'
                else:
                    data.loc[i, 'Стоимость'] = ranges[j]
                break

def date_att_removing_and_local_gen(data: pd.DataFrame):
    data.drop('Дата приезда', inplace=True, axis=1)
    for i in data.index:
        date_val = datetime.strptime(data.loc[i, 'Дата отъезда'], '%Y-%m-%dT%H:%M')
        if date_val.month in [12, 1, 2]:
            data.loc[i, 'Дата отъезда'] = 'Зима'
        elif date_val.month in [3, 4, 5]:
            data.loc[i, 'Дата отъезда'] = 'Весна'
        elif date_val.month in [6, 7, 8]:
            data.loc[i, 'Дата отъезда'] = 'Лето'
        elif date_val.month in [9, 10, 11]:
            data.loc[i, 'Дата отъезда'] = 'Осень'
    data.rename(columns={'Дата отъезда': 'Приезд / Отъезд'}, inplace=True)

def remove_fio(data: pd.DataFrame):
    data.drop('ФИО', inplace=True, axis=1)

def remove_seats(data: pd.DataFrame):
    data.drop('Вагон и место', inplace=True, axis=1)

def remove_passport(data: pd.DataFrame):
    data.drop('Паспортные данные', inplace=True, axis=1)

def remove_fromto(data: pd.DataFrame):
    data.drop("Откуда", inplace=True, axis=1)
    data.drop("Куда", inplace=True, axis=1)

def create_sheets_with_columns():
    global df, processed_file_path

    # Create a new Excel writer
    with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
        # Create two new DataFrames for the additional sheets
        sheet1_data = df[['Приезд / Отъезд', 'Рейс', 'Стоимость']]
        sheet2_data = df[['Приезд / Отъезд', 'Карта оплаты', 'Банк']]

        # Write the additional sheets to the Excel file
        sheet1_data.to_excel(writer, index=False, sheet_name='sheet_1')
        sheet2_data.to_excel(writer, index=False, sheet_name='sheet_2')



def apply_operations():
    global df, processed_file_path, listbox
    if df is not None:
        card_masking(df)
        date_att_removing_and_local_gen(df)
        remove_fio(df)
        remove_seats(df)
        remove_passport(df)
        remove_fromto(df)
        price_range(df)
        race_range(df)
        df = df.sample(frac=1).reset_index(drop=True)
        processed_file_path = file_path.replace(".xlsx", "_processed.xlsx")
        create_sheets_with_columns()
        subprocess.Popen(processed_file_path, shell=True)
        listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)
        for column in df.columns:
            listbox.insert(tk.END, column)
        listbox.pack(pady=10)
        k_anon_button.pack(pady=10)

def calculate_k_anonimity():
    global df
    selected_columns = [listbox.get(idx) for idx in listbox.curselection()]
    result = k_anonimity(df, selected_columns)
    
    if result:
        last_five_k_values = result[-5:]  # Получаем последние 5 значений
        result_label.config(text=f"Last 5 K-Anonymity Values: {last_five_k_values}")
    else:
        result_label.config(text="No K-Anonymity values found.")

def open_file():
    global file_path, df
    file_path = filedialog.askopenfilename()
    df = pd.read_excel(file_path, engine='openpyxl', index_col='Unnamed: 0')
    apply_button.pack(pady=10)

# Creating the GUI
root = tk.Tk()
root.title("Data Anonymization Tool")

file_button = tk.Button(root, text="Open File", command=open_file)
file_button.pack(pady=10)

apply_button = tk.Button(root, text="Apply Operations", command=apply_operations)

result_label = tk.Label(root, text="")
result_label.pack(pady=10)

k_anon_button = tk.Button(root, text="Calculate K-Anonymity", command=calculate_k_anonimity)

root.mainloop()