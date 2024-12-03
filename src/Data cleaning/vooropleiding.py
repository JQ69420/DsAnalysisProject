import pandas as pd
import os

def vooropleiding():
    # File paths
    students = os.path.join("..", "..", "data", "processed", "combined_students.xlsx")
    master_list = os.path.join("..", "..", "data", "processed", "Merged_Student_Data_Schoollevel_final.xlsx")

    # Load data into DataFrames
    df_students = pd.read_excel(students)
    df_master_list = pd.read_excel(master_list)

    # Assuming 'student' is the column for IDs in df_students and 'id' in df_master_list
    common_ids = set(df_master_list['id'])  # Extract IDs from master list
    df_students_filtered = df_students[df_students['id'].isin(common_ids)]  # Filter students
    unique_count = len(df_students_filtered['id'].unique())
    print(f"Number of unique strings: {unique_count}")
    


    master_list_ids = set(df_master_list['id'])
    merged_ids = set(df_students_filtered['id'])
    missing_ids = master_list_ids - merged_ids

    if missing_ids:
        print(f"Missing student IDs: {missing_ids}")
    else:
        print("No missing students.")
    
vooropleiding()
