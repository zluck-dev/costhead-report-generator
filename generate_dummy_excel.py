
import pandas as pd
import random

# Generate sample student data again
names = [f"Student_{i}" for i in range(1, 101)]  # 100 students
subjects = ["Math", "Science", "English", "History", "Computer"]

# Random marks for each student
data = {"Name": names}
for subj in subjects:
    data[subj] = [random.randint(40, 100) for _ in range(len(names))]

# Create DataFrame
df = pd.DataFrame(data)
df["Total"] = df[subjects].sum(axis=1)

# Save to Excel
file_path = "students.xlsx"
df.to_excel(file_path, index=False)

file_path