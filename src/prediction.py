import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
import os

# Load the data
student_file = os.path.join("..", "data", "processed", "Merged_Final_File_Updated.xlsx")
df = pd.read_excel(student_file)

# Map dependent variable 'dropped out' to binary
df['dropped out'] = df['dropped out'].map({'no': 0, 'yes': 1})

# Define features and target
features = ['anl1 final grade', 'anl2 final grade', 'anl3 final grade', 'anl4 final grade', 'education_level']
target = 'dropped out'

X = df[features]
y = df[target]

# Define preprocessing for numerical and categorical features
numerical_features = ['anl1 final grade', 'anl2 final grade', 'anl3 final grade', 'anl4 final grade']
categorical_features = ['education_level']

numerical_transformer = Pipeline(steps=[
    ('imputer', SimpleImputer(strategy='constant', fill_value=1)),  # Fill NA values with 1
    ('scaler', StandardScaler())
])
categorical_transformer = OneHotEncoder(handle_unknown='ignore')

preprocessor = ColumnTransformer(
    transformers=[
        ('num', numerical_transformer, numerical_features),
        ('cat', categorical_transformer, categorical_features)
    ]
)

# Create the pipeline with Gradient Boosting Classifier
pipeline = Pipeline(steps=[
    ('preprocessor', preprocessor),
    ('classifier', GradientBoostingClassifier(n_estimators=50, learning_rate=0.05, max_depth=3, random_state=42))
])

# Split data into training and testing sets for model training
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)
pipeline.fit(X_train, y_train)

# Console application
def console_app():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("Student Dropout Prediction")
    print("==========================")
    try:
        # Ask user if they want to predict seen or unseen data
        mode = input("Do you want to predict 'seen' or 'unseen' data? (type 'seen' or 'unseen'): ").strip().lower()
        if mode not in ['seen', 'unseen']:
            print("Invalid input! Please type 'seen' or 'unseen'.")
            return

        # Gather input from the user
        anl1_grade = float(input("Enter ANL 1 final grade: "))
        anl2_grade = float(input("Enter ANL 2 final grade: "))
        anl3_grade = float(input("Enter ANL 3 final grade: "))
        anl4_grade = float(input("Enter ANL 4 final grade: "))
        education_level = input("Enter education level (e.g., HAVO, VWO, MBO): ").strip().lower()

        # Create input dataframe for prediction
        user_input = pd.DataFrame({
            'anl1 final grade': [anl1_grade],
            'anl2 final grade': [anl2_grade],
            'anl3 final grade': [anl3_grade],
            'anl4 final grade': [anl4_grade],
            'education_level': [education_level]
        })

        # Make a prediction
        prediction = pipeline.predict(user_input)
        prediction_label = "Yes" if prediction[0] == 1 else "No"

        if mode == 'seen':
            # Check the actual value in the dataset
            mask = (
                ((df['anl1 final grade'].isna()) & (anl1_grade == 1) | (df['anl1 final grade'] == anl1_grade)) &
                ((df['anl2 final grade'].isna()) & (anl2_grade == 1) | (df['anl2 final grade'] == anl2_grade)) &
                ((df['anl3 final grade'].isna()) & (anl3_grade == 1) | (df['anl3 final grade'] == anl3_grade)) &
                ((df['anl4 final grade'].isna()) & (anl4_grade == 1) | (df['anl4 final grade'] == anl4_grade)) &
                (df['education_level'] == education_level)
            )

            matching_row = df[mask]
            if not matching_row.empty:
                actual_value = "Yes" if matching_row['dropped out'].values[0] == 1 else "No"
                print(f"\nPrediction: Will the student drop out? {prediction_label}")
                print(f"Actual Value in Dataset: {actual_value}")
            else:
                print("No matching record found in the dataset.")
        else:
            print(f"\nPrediction: {prediction_label}")

    except ValueError:
        print("Invalid input! Please enter the correct values.")

# Run the console application
if __name__ == "__main__":
    console_app()
