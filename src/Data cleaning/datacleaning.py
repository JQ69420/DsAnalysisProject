import pandas as pd
import os
import matplotlib.pyplot as plt

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

def add_dropout_date():
    # File paths
    master_student_list = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_final.xlsx")
    combined_students = os.path.join("..", "..", "data", "processed", "combined_students_unique.xlsx")

    # Load and preprocess the data
    df_master = pd.read_excel(master_student_list)
    df_master.columns = df_master.columns.str.lower()

    df_combined_students = pd.read_excel(combined_students)
    df_combined_students.columns = df_combined_students.columns.str.lower()

    # Ensure 'dropped out' column exists in df_master
    if 'dropped out' not in df_master.columns:
        raise ValueError("The master student list must have a 'dropped out' column.")

    # Initialize a new column for dropout dates
    df_master['dropout date'] = ''

    # Loop through each student with 'dropped out' = 'yes'
    for index, row in df_master.iterrows():
        if row['dropped out'].strip().lower() == 'yes':
            student_id = row['id']

            # Filter the combined_students DataFrame for matching IDs
            matched_rows = df_combined_students[df_combined_students['id'] == student_id]

            if not matched_rows.empty:
                # Get unique dates from stakingsdatum
                unique_dates = matched_rows['stakingsdatum'].dropna().unique()

                # If multiple dates exist, log a warning and pick the first one
                if len(unique_dates) > 1:
                    print(f"Warning: Multiple dropout dates found for ID {student_id}. Using the first one.")
                dropout_date = unique_dates[0] if unique_dates.size > 0 else 'unknown'
            else:
                dropout_date = 'unknown'

            # Assign the dropout date
            df_master.at[index, 'dropout date'] = dropout_date

    # Save the updated master student list back to an Excel file
    df_master.to_excel(master_student_list, index=False)

# Run the function
# add_dropout_date()


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

# End_Result_ANL3()

def update_outcomes():
    # Load the master student list
    master_student_list = os.path.join("..", "..", "data", "processed", "Cleaned_master_student_list_final.xlsx")
    df = pd.read_excel(master_student_list)

    # Identify columns with grades and outcomes
    grade_columns = [col for col in df.columns if "grade" in col.lower()]
    outcome_columns = [col for col in df.columns if "outcome" in col.lower()]

    # Update the outcomes based on the grade thresholds
    for grade_col, outcome_col in zip(grade_columns, outcome_columns):
        # Ensure the grades are treated as floats
        df[grade_col] = pd.to_numeric(df[grade_col], errors='coerce')

    # Identify columns with grades and outcomes
    grade_columns = [col for col in df.columns if "grade" in col.lower()]
    outcome_columns = [col for col in df.columns if "outcome" in col.lower()]

    # Update the outcomes based on the grade thresholds
    for grade_col, outcome_col in zip(grade_columns, outcome_columns):
        # Ensure the grades are treated as floats
        df[grade_col] = pd.to_numeric(df[grade_col], errors='coerce')

        # Define the conditions and corresponding outcomes
        conditions = [
            (df[grade_col] >= 0) & (df[grade_col] < 3),
            (df[grade_col] >= 3) & (df[grade_col] < 5.5),
            (df[grade_col] >= 5.5) & (df[grade_col] < 7.5),
            (df[grade_col] >= 7.5) & (df[grade_col] <= 10),
        ]
        outcomes = [
            "FAIL MISERABLY",
            "FAIL",
            "PASS",
            "PASS GREATLY",
        ]

        # Update the outcome column only for rows where the grade is not NaN
        df.loc[df[grade_col].notna(), outcome_col] = pd.cut(
            df.loc[df[grade_col].notna(), grade_col],
            bins=[-float("inf"), 3, 5.5, 7.5, float("inf")],
            labels=outcomes,
            include_lowest=True,
        ).astype(str)

    # Save the updated DataFrame back to the same file
    df.to_excel(master_student_list, index=False)
    print(f"Outcomes updated and saved back to {master_student_list}")

# Use the function on your file
# update_outcomes()



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

# End_Result_ANL4()

def male_female_dropout():
    master_file = os.path.join("..", "..", "data", "processed", 'Cleaned_master_student_list_final.xlsx')
    attendance_file = os.path.join("..", "..", "data", "raw", 'ANL1-2020-2021 attendance.xlsx')

    df_master = pd.read_excel(master_file)
    df_attendance = pd.read_excel(attendance_file)

    df_combined = pd.merge(df_master, df_attendance, on='id')

    df_filtered = df_combined[df_combined['Gender'].isin(['M', 'V'])]

    dropout_rates = (
        df_filtered.groupby('Gender')['dropped out']
        .apply(lambda x: (x == 'yes').sum() / len(x) * 100)
        .reset_index(name='Dropout Percentage')
    )

    dropout_rates['Gender'] = dropout_rates['Gender'].map({'M': 'Male', 'V': 'Female'})

    colors = ['blue' if gender == 'Male' else 'red' for gender in dropout_rates['Gender']]

    plt.bar(dropout_rates['Gender'], dropout_rates['Dropout Percentage'], color=colors)
    plt.xlabel('Gender')
    plt.ylabel('Dropout Percentage')
    plt.title('Dropout Percentage by Gender')
    plt.show()

# male_female_dropout()

def voltijd_deeltijd_dropout():
    # File paths
    master_file = os.path.join("..", "..", "data", "processed", 'Cleaned_master_student_list_final.xlsx')
    combined_file = os.path.join("..", "..", "data", "processed", 'combined_students_unique.xlsx')

    # Load the datasets
    df_master = pd.read_excel(master_file)
    df_combined = pd.read_excel(combined_file)

    # Merge the datasets on 'id'
    df_merged = pd.merge(df_combined, df_master, on='id')

    # Filter data for valid 'voltijd deeltijd' values (VT and DT only)
    df_filtered = df_merged[df_merged['voltijd deeltijd'].isin(['VT', 'DT'])]

    # Calculate dropout rates
    dropout_rates = (
        df_filtered.groupby('voltijd deeltijd')['dropped out']
        .apply(lambda x: (x == 'yes').sum() / len(x) * 100)
        .reset_index(name='Dropout Percentage')
    )

    # Map VT and DT to descriptive labels
    dropout_rates['voltijd deeltijd'] = dropout_rates['voltijd deeltijd'].map({'VT': 'Voltijd', 'DT': 'Deeltijd'})

    # Define colors for bars
    colors = ['green' if vt_dt == 'Voltijd' else 'orange' for vt_dt in dropout_rates['voltijd deeltijd']]

    # Plot the bar chart
    plt.bar(dropout_rates['voltijd deeltijd'], dropout_rates['Dropout Percentage'], color=colors)
    plt.xlabel('Type of Enrollment')
    plt.ylabel('Dropout Percentage')
    plt.title('Dropout Percentage by Enrollment Type')
    plt.show()

# voltijd_deeltijd_dropout()

def dropout_by_outcome():
    master_file = os.path.join("..", "..", "data", "processed", 'Cleaned_master_student_list_final.xlsx')
    df_master = pd.read_excel(master_file)

    # Identify columns that end with 'outcome'
    outcome_columns = [col for col in df_master.columns if col.endswith('outcome')]

    # Define the fixed order and corresponding colors
    outcome_order = ['FAIL MISERABLY', 'FAIL', 'PASS', 'PASS GREATLY']
    outcome_colors = {'FAIL MISERABLY': 'red', 'FAIL': 'orange', 'PASS': 'lightgreen', 'PASS GREATLY': 'darkgreen'}

    # Loop through each 'outcome' column and calculate dropout percentages
    for outcome_col in outcome_columns:
        # Group by outcome column and calculate dropout percentage
        dropout_rates = (
            df_master.groupby(outcome_col)['dropped out']
            .apply(lambda x: (x == 'yes').sum() / len(x) * 100)
            .reset_index(name='Dropout Percentage')
        )

        # Reorder the data to match the fixed order
        dropout_rates[outcome_col] = pd.Categorical(
            dropout_rates[outcome_col], categories=outcome_order, ordered=True
        )
        dropout_rates = dropout_rates.sort_values(outcome_col)

        # Plot the chart for this outcome column
        plt.figure(figsize=(8, 6))
        plt.bar(
            dropout_rates[outcome_col],
            dropout_rates['Dropout Percentage'],
            color=[outcome_colors[outcome] for outcome in dropout_rates[outcome_col]],
        )
        plt.xlabel(outcome_col.replace('_', ' ').title())
        plt.ylabel('Dropout Percentage')
        plt.title(f'Dropout Percentage by {outcome_col.replace("_", " ").title()}')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

# dropout_by_outcome()

def dropout_by_attendance():
    file_path = os.path.join("..", "..", "data", "processed", 'students_with_attendance_and_homework.xlsx')
    df = pd.read_excel(file_path)

    # Replace empty values in attendance columns with 0
    df['ANL1 Attendance'] = df['ANL1 Attendance'].fillna(0)
    df['ANL2 Attendance'] = df['ANL2 Attendance'].fillna(0)

    # Define attendance columns
    attendance_columns = ['ANL1 Attendance', 'ANL2 Attendance']

    # Function to bin attendance into groups
    def bin_attendance(value):
        return f"{int(value // 10) * 10}-{int(value // 10) * 10 + 10}"

    # Loop through attendance columns
    for col in attendance_columns:
        # Create a new column for binned attendance
        df[f'{col} Group'] = df[col].apply(bin_attendance)

        # Calculate dropout percentages for each group
        dropout_rates = (
            df.groupby(f'{col} Group')['dropped out']
            .apply(lambda x: (x == 'yes').sum() / len(x) * 100)
            .reset_index(name='Dropout Percentage')
        )

        # Sort the groups in the correct order
        dropout_rates = dropout_rates.sort_values(f'{col} Group', key=lambda x: [int(g.split('-')[0]) for g in x])

        # Plot the bar chart
        plt.figure(figsize=(10, 6))
        plt.bar(dropout_rates[f'{col} Group'], dropout_rates['Dropout Percentage'], color='skyblue')
        plt.xlabel(f'{col} (Grouped)')
        plt.ylabel('Dropout Percentage')
        plt.title(f'Dropout Percentage by {col}')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

# dropout_by_attendance()


