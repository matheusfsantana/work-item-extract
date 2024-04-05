import os
from datetime import datetime
import pandas as pd
import warnings

warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings('ignore', category=FutureWarning)


def return_csv_file():
    # File directory
    csv_file_dir = 'upload_csv'

    # Read the files inside the directory
    files = os.listdir(csv_file_dir)

    # Read all csv files inside the directory
    csv_files = [file for file in files if file.endswith('.csv')]

    if len(csv_files) == 1:  
        return os.path.join('upload_csv', csv_files[0]) 
    elif len(csv_files) > 1:
        print("There is more than one CSV file in the directory. Unable to determine which one to read.")
        return None
    else:
        print("There is no CSV file in the directory.")
        return None

def work_item_extact():

    csv_file = return_csv_file()

    if csv_file is None:
        return None
    
    print("Treating the CSV file...")

    df = pd.read_csv(csv_file, delimiter=',')
    df = df.rename(columns={"Iteration Level 4":'Sprint'})
    
    # Separating User Story from other Work Itens Type
    df_user_story = df[df['Work Item Type'] == 'User Story']
    df_work_itens = df[df['Work Item Type'] != 'User Story']

    for index, row in df_work_itens.iterrows():

        # If its empty, Sprint = 0
        if pd.isna(df_work_itens.at[index, 'Sprint']) or None:
            df_work_itens.at[index, 'Sprint'] = int(0)
        else:
            # If its not empty, get the sprint number
            df_work_itens.at[index, 'Sprint'] = str(row['Sprint']).split('Sprint', 1)[1].strip()
            df_work_itens.at[index, 'Sprint'] = int(df_work_itens.at[index, 'Sprint'])

        # Get only the worker name
        df_work_itens.at[index, 'Tester'] = str(row['Tester']).split('<', 1)[0].strip()
        df_work_itens.at[index, 'Assigned To'] = str(row['Assigned To']).split('<', 1)[0].strip()

        # Get only the date
        df_work_itens.at[index, 'Due Date'] = str(row['Due Date']).split(' ', 1)[0].strip()

        # Get the severity name without numbers 
        # Ex: Severity: 1 - HIGH -> Severity: HIGH
        if '-' in str(df_work_itens.at[index, 'Severity']):
            df_work_itens.at[index, 'Severity'] = str(df_work_itens.at[index, 'Severity']).split(' - ')[1]
        else:
            df_work_itens.at[index, 'Severity'] = '-'

        # Associating the User Story with a Task/Bug
        matching_row = df_user_story[df_user_story['ID'] == row['Parent']]

        if not matching_row.empty:
            df_work_itens.at[index, 'User Story'] = matching_row.iloc[0]['Title']
        else:
            df_work_itens.at[index, 'User Story'] = None 

        # If
        for col, value in row.items():
            if pd.isna(value) or value == '':
                df_work_itens.at[index, col] = '-'
    
    # Ordening the columns
    columns_new_order = ['ID', 'Sprint', 'User Story','Work Item Type', 'Title', 'Assigned To', 'Tester', 'Ambiente', 'Severity', 'Tags', 'State', 'Created Date', 'Due Date', 'Activated Date','Closed Date']
    df_work_itens = df_work_itens[columns_new_order]

    return df_work_itens


def return_excel_file():

    df = work_item_extact()
    if df is None:
        return None
    
    print("Writing to XLSX...")
    
    filename = f'xlsx_file/{datetime.now().strftime("%Y%m%d%H%M%S")}-work_item.xlsx'

    # Writing the extracted dataframe to excel
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Work Itens')
    worksheet = writer.sheets['Work Itens']
    worksheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)  # Adding filters for the sheet

    writer.close()

    print("File saved in:")
    print(os.path.join(os.getcwd(), filename))

def main():
    return_excel_file()


if __name__ == "__main__":
   main()