import pandas as pd
import os

master_student_list = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list.xlsx")
combined_students = os.path.join("..", "..", "data", "processed", "combined_students.xlsx")

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

def add_dropout_column():

    pd.set_option('display.max_rows', None)
    
    df_master = pd.read_excel(master_student_list)
    df_master.columns = df_master.columns.str.lower()
    
    df_combined_students = pd.read_excel(combined_students)
    df_combined_students.columns = df_combined_students.columns.str.lower()
    
    filter_condition = lambda df: (df['school'] == 'Hogeschool Rotterdam') & (df['opleiding'].str.startswith('INF'))
    df_students_filtered = df_combined_students[filter_condition(df_combined_students)]
    
    merged_df = df_master.merge(df_students_filtered[['id', 'stakingsdatum', 'school', 'opleiding']], on='id', how='left')
    
    # Group by ID and determine dropout status based on 'Stakingsdatum' presence
    # yes = all rows of stakingsdatum have a date
    # no = no rows of stakingsdatum have a date
    # ? = some rows of stakingsdatum have a date and some rows do not
    dropout_status = merged_df.groupby('id')['stakingsdatum'].apply(
        lambda x: 'yes' if x.notna().all() else 'no'
    )
    
    df_master['dropped out'] = df_master['id'].map(dropout_status)

    df_master.to_excel(master_student_list, index=False)

# add_dropout_column()


def End_Result_ANL3():
    ANL_3_results = os.path.join("..", "..", "data", "processed", "ANL3_FC&SC_Student_Merge.xlsx")

    df_ANL3 = pd.read_excel(ANL_3_results)


    df_ANL3['ANL3 Final Grade'] = df_ANL3[['ANL3 Fc Grade', 'ANL3 Sc Grade']].max(axis=1)

    
    df_ANL3['ANL3 Final Result'] = df_ANL3.apply(
    lambda row: 'PASS' if row['ANL3 Fc Outcome'] == 'PASS' or row['ANL3 Sc Outcome'] == 'PASS' else 'FAIL',
    axis=1
    )

    output_file = os.path.join("..", "..", "data", "processed", "ANL3_FULLY_MERGED.xlsx")
    df_ANL3.to_excel(output_file, index=False)



#End_Result_ANL3()


def End_Result_ANL1():
    ANL_1_results = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_final.xlsx")

    df_ANL1 = pd.read_excel(ANL_1_results)


    df_ANL1['ANL1 Final Grade'] = df_ANL1[['ANL1 Fc Grade', 'ANL1 Sc Grade']].max(axis=1)

    
    df_ANL1['ANL1 Final Result'] = df_ANL1.apply(
    lambda row: 'PASS' if row['ANL1 Fc Outcome'] == 'PASS' or row['ANL1 Sc Outcome'] == 'PASS' else 'FAIL',
    axis=1
    )

    output_file = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_finalv2.xlsx")
    df_ANL1.to_excel(output_file, index=False)



def End_Result_ANL2():
    ANL_1_results = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_finalv2.xlsx")

    df_ANL1 = pd.read_excel(ANL_1_results)


    df_ANL1['ANL2 Final Grade'] = df_ANL1[['ANL2 Fc Grade', 'ANL2 Sc Grade']].max(axis=1)

    
    df_ANL1['ANL2 Final Result'] = df_ANL1.apply(
    lambda row: 'PASS' if row['ANL2 Fc Outcome'] == 'PASS' or row['ANL2 Sc Outcome'] == 'PASS' else 'FAIL',
    axis=1
    )

    output_file = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_finalv3.xlsx")
    df_ANL1.to_excel(output_file, index=False)


def End_Result_ANL4():
    ANL_1_results = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_finalv3.xlsx")

    df_ANL1 = pd.read_excel(ANL_1_results)


    df_ANL1['ANL4 Final Grade'] = df_ANL1[['ANL4 Fc Grade', 'ANL4 Sc Grade']].max(axis=1)

    
    df_ANL1['ANL4 Final Result'] = df_ANL1.apply(
    lambda row: 'PASS' if row['ANL4 Fc Outcome'] == 'PASS' or row['ANL4 Sc Outcome'] == 'PASS' else 'FAIL',
    axis=1
    )

    output_file = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_finalv4.xlsx")
    df_ANL1.to_excel(output_file, index=False)

End_Result_ANL4()