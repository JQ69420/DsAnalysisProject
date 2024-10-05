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

Drop_Duplicated()