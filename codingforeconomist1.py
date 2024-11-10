
"""
@author: Nihad Eldarov
"""
from scipy.stats import zscore
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Reading the Excel file
FitnessMembersData = pd.read_excel(io='C:/Users/Nihad/Downloads/Fitness_Members_Data.xlsx')

# Display the first five rows
print(FitnessMembersData.head())

FitnessMembersData['Age_Group'] = pd.cut(FitnessMembersData['Age'], bins=[0, 25, 40, 60, 100], labels=['Young', 'Middle', 'Middle-Aged', 'Elderly'])

# Show the count of missing values
print("Count of missing values in the Monthly_Calories_Burned column:")
print(FitnessMembersData['Monthly_Calories_Burned'].isnull().sum())  
print("Count of missing values in the Avg_Exercise_Duration column:")
print(FitnessMembersData['Avg_Exercise_Duration'].isnull().sum())  

# Imputing missing values with median
FitnessMembersData['Avg_Exercise_Duration'] = FitnessMembersData['Avg_Exercise_Duration'].fillna(FitnessMembersData['Avg_Exercise_Duration'].median())
# Imputing missing values with mean
FitnessMembersData['Monthly_Calories_Burned'] = FitnessMembersData['Monthly_Calories_Burned'].fillna(FitnessMembersData['Monthly_Calories_Burned'].mean())

# Show the count of missing values in the updated table
print("Count of missing values in the updated Monthly_Calories_Burned column:")
print(FitnessMembersData['Monthly_Calories_Burned'].isnull().sum())
print("Count of missing values in the updated Avg_Exercise_Duration column:")
print(FitnessMembersData['Avg_Exercise_Duration'].isnull().sum())


# Calculate Z-score to identify outliers
FitnessMembersData['Z_Score_Calories'] = zscore(FitnessMembersData['Monthly_Calories_Burned'])
outliers = FitnessMembersData[FitnessMembersData['Z_Score_Calories'].abs() > 1.5]

# Z-score plot: Highlight outliers with a different color
plt.figure(figsize=(10, 6))
plt.scatter(FitnessMembersData.index, FitnessMembersData['Monthly_Calories_Burned'], label="Normal Values", color="blue")
plt.scatter(outliers.index, outliers['Monthly_Calories_Burned'], label="Outliers (Z > 1.5)", color="red")
plt.xlabel("Index")
plt.ylabel("Monthly Calories Burned")
plt.title("Outliers Visualized by Z-Score")
plt.legend()
plt.show()
# In the chart, a threshold of 1.5 Z-score was chosen to define those values that exceed 1.5 standard deviations from the mean as outliers.


# Average Calories Burned by Gender


calories_compare_gender = FitnessMembersData.groupby('Gender')['Monthly_Calories_Burned'].mean()
print("Average Calories Burned by Gender:")
print(calories_compare_gender)


# Average exercise duration by exercise type
exercise_duration = FitnessMembersData.groupby('Favorite_Exercise_Type')['Avg_Exercise_Duration'].mean()
print("Average Exercise Duration by Favorite Exercise Type:")
print(exercise_duration)

# Plotting age distribution
plt.figure(figsize=(8, 5))
FitnessMembersData['Age'].hist()
plt.title("Age Distribution")
plt.xlabel("Age")
plt.ylabel("Frequency")
plt.show()

# Plotting average calories burned by favorite exercise type
plt.figure(figsize=(10, 6))
sns.barplot(data=FitnessMembersData, x='Favorite_Exercise_Type', y='Monthly_Calories_Burned')
plt.title("Average Calories Burned by Favorite Exercise Type")
plt.show()

# Average Calories Burned by Gender
calories_compare_gender = FitnessMembersData.groupby('Gender')['Monthly_Calories_Burned'].mean().reset_index()

# Average Exercise Duration by Exercise Type
exercise_duration = FitnessMembersData.groupby('Favorite_Exercise_Type')['Avg_Exercise_Duration'].mean().reset_index()

# Saving the results and original data in the same Excel file
with pd.ExcelWriter("updated5_Fitness_Members_Data_with_Analysis.xlsx") as writer:
    FitnessMembersData.to_excel(writer, sheet_name="Original Data", index=False)
    calories_compare_gender.to_excel(writer, sheet_name="Calories by Gender", index=False)
    exercise_duration.to_excel(writer, sheet_name="Exercise Duration by Type", index=False)