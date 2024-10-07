import pandas as pd
import os

def CombineFiles():
    students_file_2 = os.path.join("..", "..", "data", "raw", "(anonymised) Eerstejaars studenten INF met einddatum en reden - additional students TWO.xlsx")
    students_file_1 = os.path.join("..", "..", "data", "raw", "(anonymised) Eerstejaars studenten INF met einddatum en reden ONE.xlsx")


    # Read the Excel files directly
    df_student_1 = pd.read_excel(students_file_1)  # For reading Excel files
    df_student_2 = pd.read_excel(students_file_2)

    df_student_1.columns = df_student_1.columns.str.strip().str.lower()
    df_student_2.columns = df_student_2.columns.str.strip().str.lower()


    df_combined = pd.concat([df_student_1, df_student_2], ignore_index=True)
    output_file = os.path.join("..", "..", "data", "processed", "combined_students.xlsx")
    df_combined.to_excel(output_file, index=False)  # index=False to avoid writing row indices
    if os.path.exists(output_file):
        print(f"The file '{output_file}' already exists.")
    else:
        # Save the combined DataFrame to a new Excel file if it doesn't exist
        df_combined.to_excel(output_file, index=False)  # index=False to avoid writing row indices
        print(f"Combined data saved to {output_file}")
    

def Drop_Duplicated():
    student_file = os.path.join("..", "..", "data", "processed", "combined_students.xlsx")
    df_students = pd.read_excel(student_file)
    df_students_unique = df_students.drop_duplicates()

    output_file = os.path.join("..", "..", "data", "processed", "combined_students_unique.xlsx")
    df_students_unique.to_excel(output_file, index=False)
    print(f"Unique rows saved to: {output_file}")

def Merge_Files_ANL2_Attandace():
    attandace_file = os.path.join("..", "..", "data", "raw", "ANL2-2020-2021 Homework and Attendance.xlsx")
    student_file = os.path.join("..", "..", "data", "processed", "combined_students.xlsx")

    df_students = pd.read_excel(student_file)
    df_attandace = pd.read_excel(attandace_file)

    df_students['id'] = df_students['id'].fillna(0).astype(int)
    df_attandace['id'] = df_attandace['id'].fillna(0).astype(int)  

    df_students['id'] = df_students['id'].astype(str).str.strip()
    df_attandace['id'] = df_attandace['id'].astype(str).str.strip()

    print(df_attandace.head)
    merged_data = pd.merge(df_students, df_attandace[['id', 'Attendance %', "Team assignment"]], on='id', how='left')
    merged_data.rename(columns={'Attendance %': 'Attandace ANL2', 'Team assignment': 'Team Assignment ANL2'}, inplace=True)  # Change column name
    column_order = list(df_students.columns) + ['Attandace ANL2' ,  'Team Assignment ANL2'] 
    merged_data = merged_data[column_order]

    output_file = os.path.join("..", "..", "data", "processed", "Merged_Attandace_ANL2.xlsx")
    merged_data.to_excel(output_file, index=False)



def Merge_Files_ANL1_Attandace(): # werkt nog niet
    attandace_file = os.path.join("..", "..", "data", "raw", "ANL1-2020-2021 attendance.xlsx")
    student_file = os.path.join("..", "..", "data", "processed", "Merged_Attandace_ANL2.xlsx")

    df_students = pd.read_excel(student_file)
    df_attandace = pd.read_excel(attandace_file)

    df_students['id'] = df_students['id'].fillna(0).astype(int)
    df_attandace['id'] = df_attandace['id'].fillna(0).astype(int)  

    df_students['id'] = df_students['id'].astype(str).str.strip()
    df_attandace['id'] = df_attandace['id'].astype(str).str.strip()

    for s in df_students['id']:
        for s1 ,s2 in zip(df_attandace['id'], df_attandace['Attendance %']):
            if(s == s1):
                print("TRUE")
                print(s2)
    print(df_attandace.head)
    merged_data = pd.merge(df_students, df_attandace[['id', 'Attendance %', "Team assignment"]], on='id', how='left')
    merged_data.rename(columns={'Attendance %': 'Attandace ANL1', 'Team assignment': 'Team Assignment ANL1'}, inplace=True)  # Change column name
    column_order = list(df_students.columns) + ['Attandace ANL1' ,  'Team Assignment ANL1'] 
    merged_data = merged_data[column_order]

    output_file = os.path.join("..", "..", "data", "processed", "Merged_students_ANL1+ANL2.xlsx")
    merged_data.to_excel(output_file, index=False)
    


def Clean_Master_list():
    students = os.path.join("..", "..", "data", "raw", "Master list for students.xlsx")

    df_students = pd.read_excel(students)

    filtered_df = df_students[df_students['Group'].apply(lambda x: len(str(x)) == 1) | (df_students['Group'] == 'DINF1') | (df_students['Group'] == 'DINF2')]

    sorted_df = filtered_df.sort_values(by='Group')

    output_path = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list.xlsx")
    sorted_df.to_excel(output_path, index=False)

def Merge_ANL3_FC():
    ANL_3_results = os.path.join("..", "..", "data", "raw", "INFANL3-2020-2021 EXAM first chance.xlsx")
    student_file = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list.xlsx")

    df_students = pd.read_excel(student_file)
    df_ANL_3_results= pd.read_excel(ANL_3_results , sheet_name="Grades")

    merged_data = pd.merge(df_students, df_ANL_3_results[["ID", "Grade", "Outcome"]], on='ID', how='left')
    merged_data.rename(columns={'Grade': 'ANL3 Fc Grade', 'Outcome': 'ANL3 Fc Outcome'}, inplace=True)  # Change column name
    column_order = list(df_students.columns) + ['ANL3 Fc Grade' ,  'ANL3 Fc Outcome'] 
    merged_data = merged_data[column_order]

    output_file = os.path.join("..", "..", "data", "processed", "ANL3_FC_Student_Merge.xlsx")
    merged_data.to_excel(output_file, index=False)


def Merge_ANL3_SC():
    ANL_3_results_SC = os.path.join("..", "..", "data", "raw", "INFANL3-2020-2021 EXAM second chance.xlsx")
    student_file = os.path.join("..", "..", "data", "processed", "ANL3_FC_Student_Merge.xlsx")

    df_students = pd.read_excel(student_file)
    df_ANL_3_results= pd.read_excel(ANL_3_results_SC , sheet_name="Grades")

    merged_data = pd.merge(df_students, df_ANL_3_results[["ID", "Grade", "Outcome"]], on='ID', how='left')
    merged_data.rename(columns={'Grade': 'ANL3 Sc Grade', 'Outcome': 'ANL3 Sc Outcome'}, inplace=True)  # Change column name
    column_order = list(df_students.columns) + ['ANL3 Sc Grade' ,  'ANL3 Sc Outcome'] 
    merged_data = merged_data[column_order]

    output_file = os.path.join("..", "..", "data", "processed", "ANL3_FC&SC_Student_Merge.xlsx")
    merged_data.to_excel(output_file, index=False)


Merge_ANL3_SC()