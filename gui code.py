import pandas as pd
import tkinter as tk
from tkinter import ttk

# Read your Excel file into a DataFrame and strip column names of leading/trailing whitespaces
df = pd.read_excel('C:\\Users\\parig\\OneDrive\\Desktop\\internship_folder\\computation data to be used.xlsx', header=2)
df.columns = df.columns.str.strip()

# Get unique campaigns in the dataset
unique_campaigns = df['Campaign'].unique()

# Create a tkinter window
window = tk.Tk()
window.title("Campaign Disposition Summary")

# Create a tabular view for the results
tabular_view = ttk.Treeview(window, columns=("Positive Disposition", "P Count", "P % Age", "P Total Call Length", "P Avg AHT", "Negative Disposition", "N Count", "N % Age", "N Total Call Length", "N Avg AHT"))
tabular_view.heading("#1", text="Positive Disposition")
tabular_view.heading("#2", text="P Count")
tabular_view.heading("#3", text="P % Age")
tabular_view.heading("#4", text="P Total Call Length")
tabular_view.heading("#5", text="P Avg AHT")
tabular_view.heading("#6", text="Negative Disposition")
tabular_view.heading("#7", text="N Count")
tabular_view.heading("#8", text="N % Age")
tabular_view.heading("#9", text="N Total Call Length")
tabular_view.heading("#10", text="N Avg AHT")

tabular_view.pack()

# Loop through each campaign
for campaign in unique_campaigns:
    campaign_df = df[df['Campaign'] == campaign]
    
    positive_dispositions = campaign_df[campaign_df['Best Disposition (P / N'] == 'Positive']
    negative_dispositions = campaign_df[campaign_df['Best Disposition (P / N'] == 'Negative']

    positive_counts = positive_dispositions['Best Disposition'].value_counts()
    positive_percentage = (positive_counts / positive_dispositions.shape[0]) * 100
    positive_total_call_length = positive_dispositions['Call Duration'].sum()

    negative_counts = negative_dispositions['Best Disposition'].value_counts()
    negative_percentage = (negative_counts / negative_dispositions.shape[0]) * 100
    negative_total_call_length = negative_dispositions['Call Duration'].sum()

    total_call_length = campaign_df['Call Duration'].sum()
    
    campaign_df['AHT'] = campaign_df['Call Duration'] / total_call_length

    for disposition in positive_counts.index:
        p_avg_aht = campaign_df[campaign_df['Best Disposition'] == disposition]['AHT'].mean()
        n_count = negative_counts.get(disposition, 0)
        n_percentage = negative_percentage.get(disposition, 0)
        n_avg_aht = campaign_df[campaign_df['Best Disposition'] == disposition]['AHT'].mean()
        tabular_view.insert("", "end", values=(disposition, positive_counts[disposition], f"{positive_percentage[disposition]:.2f}%", positive_total_call_length, f"{p_avg_aht:.2f}",
                                              disposition, n_count, f"{n_percentage:.2f}%", negative_total_call_length, f"{n_avg_aht:.2f}"))

# Start the tkinter main loop
window.mainloop()
