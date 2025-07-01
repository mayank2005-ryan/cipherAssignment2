import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import numpy as np


output_file = Path("employee_analysis_result.xlsx")    

df = pd.read_csv("employee_performance.csv")

df["JoinDate"] = pd.to_datetime(df["JoinDate"], dayfirst=True, errors="coerce")

df["Salary"]            = pd.to_numeric(df["Salary"], errors="coerce")
df["PerformanceRating"] = pd.to_numeric(df["PerformanceRating"], errors="coerce")

df_clean = df.dropna().copy()

df_clean["Tenure"] = 2025 - df_clean["JoinDate"].dt.year

bins   = [-np.inf, 50_000, 90_000, np.inf]
labels = ["Low", "Medium", "High"]
df_clean["SalaryCategory"] = pd.cut(df_clean["Salary"], bins=bins, labels=labels)

avg_salary_by_dept = (df_clean.groupby("Department", as_index=False)["Salary"].mean().rename(columns={"Salary": "AverageSalary"}))

gender_count_by_dept = (df_clean.pivot_table(index="Department",columns="Gender",values="EmployeeID",aggfunc="count",fill_value=0).reset_index())

avg_rating_by_dept = (df_clean.groupby("Department", as_index=False)["PerformanceRating"].mean().rename(columns={"PerformanceRating": "AvgPerformanceRating"}))

low_performers = df_clean[df_clean["PerformanceRating"] <= 2]


with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df_clean.to_excel(writer,           sheet_name="Cleaned_Data",           index=False)
    avg_salary_by_dept.to_excel(writer, sheet_name="avg_salary_by_dept",     index=False)
    gender_count_by_dept.to_excel(writer, sheet_name="gender_count_by_dept", index=False)
    avg_rating_by_dept.to_excel(writer, sheet_name="avg_rating_by_dept",     index=False)
    low_performers.to_excel(writer,     sheet_name="low_performers",         index=False)


plt.figure()
avg_salary_by_dept.plot(kind="bar",x="Department",y="AverageSalary",legend=False)
plt.ylabel("Average Salary")
plt.title("Average Salary by Department")
plt.tight_layout()
plt.show()


plt.figure()
df_clean["Gender"].value_counts().plot(kind="pie",autopct="%1.1f%%",ylabel="")
plt.title("Overall Gender Distribution")
plt.tight_layout()
plt.show()
