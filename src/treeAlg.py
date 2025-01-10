import pandas as pd
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import classification_report, accuracy_score
from sklearn.preprocessing import LabelEncoder
from imblearn.over_sampling import SMOTE 

# Edit the file path to the location of the dataset
file_path = 'Merged_Final_File.xlsx'
data = pd.ExcelFile(file_path)
sheet_data = data.parse("Sheet1")

# Selecting relevant columns
relevant_columns = ['anl1 fc grade', 'anl2 fc grade', 'anl3 fc grade', 'anl4 fc grade', 'education_level', 'dropped out']
data_filtered = sheet_data[relevant_columns].copy()

# Handling missing values (dropping rows with NaNs)
data_filtered.dropna(inplace=True)

# Categorical 'education_level' and target variable 'dropped out' handling
encoder = LabelEncoder()
data_filtered['education_level'] = encoder.fit_transform(data_filtered['education_level'])
data_filtered['dropped out'] = encoder.fit_transform(data_filtered['dropped out'])

# Splitting the data into features and target
X = data_filtered[['anl1 fc grade', 'anl2 fc grade', 'anl3 fc grade', 'anl4 fc grade', 'education_level']]
y = data_filtered['dropped out']

# Handling class imbalance with SMOTE
smote = SMOTE(random_state=42)
X_resampled, y_resampled = smote.fit_resample(X, y)

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(X_resampled, y_resampled, test_size=0.3, random_state=42)

# Hyperparameter tuning with GridSearchCV
param_grid = {
    'max_depth': [3, 7, 10, None],
    'min_samples_split': [2, 5, 10],
    'min_samples_leaf': [1, 2, 3, 4],
    'class_weight': ['balanced', None]
}
grid_search = GridSearchCV(
    estimator=DecisionTreeClassifier(random_state=42),
    param_grid=param_grid,
    scoring='f1_macro',
    cv=5,
    verbose=1,
    n_jobs=-1
)

# Fit the model using GridSearchCV
grid_search.fit(X_train, y_train)
best_model = grid_search.best_estimator_

# Predictions and evaluation
y_pred = best_model.predict(X_test)
accuracy = accuracy_score(y_test, y_pred)
report = classification_report(y_test, y_pred)

print(f"Best Parameters: {grid_search.best_params_}")
print(f"Accuracy: {accuracy}")
print("Classification Report:")
print(report)
