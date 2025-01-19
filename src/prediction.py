import pandas as pd
import numpy as np
from sklearn.model_selection import StratifiedKFold, cross_val_score
from sklearn.linear_model import LogisticRegression
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

# Create the pipeline
pipeline = Pipeline(steps=[
    ('preprocessor', preprocessor),
    ('classifier', LogisticRegression(max_iter=1000, random_state=42))
])

# Fit the pipeline with the data
pipeline.fit(X, y)

# Function to make a prediction based on user input
def make_prediction():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("Student Dropout Prediction")
    print("==========================")

    try:
        # Input and validation for grades
        anl1_grade = float(input("Enter ANL 1 final grade (0-10): "))
        if not 0 <= anl1_grade <= 10:
            raise ValueError("Grade must be between 0 and 10.")

        anl2_grade = float(input("Enter ANL 2 final grade (0-10): "))
        if not 0 <= anl2_grade <= 10:
            raise ValueError("Grade must be between 0 and 10.")

        anl3_grade = float(input("Enter ANL 3 final grade (0-10): "))
        if not 0 <= anl3_grade <= 10:
            raise ValueError("Grade must be between 0 and 10.")

        anl4_grade = float(input("Enter ANL 4 final grade (0-10): "))
        if not 0 <= anl4_grade <= 10:
            raise ValueError("Grade must be between 0 and 10.")
        
        education_level = input("Enter education level (e.g., HAVO, VWO, MBO): ").strip().lower()

        user_input = pd.DataFrame({
            'anl1 final grade': [anl1_grade],
            'anl2 final grade': [anl2_grade],
            'anl3 final grade': [anl3_grade],
            'anl4 final grade': [anl4_grade],
            'education_level': [education_level]
        })
        
        # Make the prediction
        prediction = pipeline.predict(user_input)
        prediction_label = "Yes" if prediction[0] == 1 else "No"

        print(f"\nPrediction: Will the student drop out? {prediction_label}")

    except ValueError as e:
        print(f"Invalid input! {e}")

# Call the make_prediction function
if __name__ == "__main__":
    make_prediction()
