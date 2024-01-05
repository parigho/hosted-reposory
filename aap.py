from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    # Code from the first Python script
    # Code from the second Python script
    df2 = pd.read_excel(r'C:\Users\parig\OneDrive\Desktop\internship_folder\data raw.xlsx', header=2)

    df2.columns = df2.columns.str.strip()

    campaign_summary = df2.groupby('Campaign').agg({
        'phone_number': 'nunique',
        'Attempts': 'sum',
        'Connect': 'sum',
        '1ST DISPOSITION (P/N)': lambda x: (x == 'Positive').sum(),
    }).reset_index()

    campaign_summary = campaign_summary.rename(columns={
        'phone_number': 'Unique Leads',
        'Attempts': 'Attempts',
        'Connect': 'Connect',
        '1ST DISPOSITION (P/N)': 'Positive'
    })

    campaign_summary['Negative'] = campaign_summary['Attempts'] - campaign_summary['Positive']
    campaign_summary['Intensity'] = round(campaign_summary['Attempts'] / campaign_summary['Unique Leads'], 2)
    campaign_summary['Connect %'] = round(campaign_summary['Connect'] / campaign_summary['Attempts'] * 100)
    campaign_summary['Positive %'] = round(campaign_summary['Positive'] / campaign_summary['Attempts'] * 100, 2)
    campaign_summary['Negative %'] = round(campaign_summary['Negative'] / campaign_summary['Attempts'] * 100, 2)
    campaign_summary['Total %'] = round(campaign_summary['Positive %'] + campaign_summary['Negative %'])

    # Save the results to variables
    result2 = campaign_summary[['Campaign', 'Unique Leads', 'Attempts', 'Intensity', 'Connect', 'Connect %', 'Positive', 'Negative', 'Positive %', 'Negative %', 'Total %']]
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Code for the third Python script
    # Replace 'path_to_your_third_data.xlsx' with the actual path to your third data file
    df3 = pd.read_excel(r"C:\Users\parig\OneDrive\Desktop\internship_folder\datatest2.xlsx", header=2)

    # Add your code for processing df3 and generating result_df
    # ...

    # Display the result in HTML
    result_df = df3.groupby('full_name')['Connect Status'].value_counts().unstack(fill_value=0).reset_index()
    result_df.columns = ['full_name', 'Connect', 'Non-Connect']
    result_df['Total Count'] = result_df['Connect'] + result_df['Non-Connect']
    result_df['Connect Percentage'] = round((result_df['Connect'] / result_df['Total Count']) * 100)
    result_df['Non-Connect Percentage'] = round((result_df['Non-Connect'] / result_df['Total Count']) * 100)
    result_df['Total Percentage'] = result_df['Connect Percentage'] + result_df['Non-Connect Percentage']

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Read your Excel file into a DataFrame and strip column names of leading/trailing whitespaces
    df = pd.read_excel('C:\\Users\\parig\\OneDrive\\Desktop\\internship_folder\\computation data to be used.xlsx', header=2)

    df.columns = df.columns.str.strip()

    # Get unique campaigns in the dataset
    unique_campaigns = df['Campaign'].unique()

    # Initialize an empty string to store the HTML table
    result1 = ""

    # Loop through each campaign
    for campaign in unique_campaigns:
        # Filter rows for the current campaign
        campaign_df = df[df['Campaign'] == campaign]

        # Filter rows with 'Positive' and 'Negative' dispositions based on 'Best Disposition (P/N)'
        positive_dispositions = campaign_df[campaign_df['Best Disposition (P / N'] == 'Positive']
        negative_dispositions = campaign_df[campaign_df['Best Disposition (P / N'] == 'Negative']

        # Calculate counts, percentages, total call length, and average AHT for positive dispositions
        positive_counts = positive_dispositions['Best Disposition'].value_counts()
        positive_percentage = (positive_counts / positive_dispositions.shape[0]) * 100
        positive_total_call_length = positive_dispositions['Call Duration'].sum()

        # Calculate counts, percentages, total call length, and average AHT for negative dispositions
        negative_counts = negative_dispositions['Best Disposition'].value_counts()
        negative_percentage = (negative_counts / negative_dispositions.shape[0]) * 100
        negative_total_call_length = negative_dispositions['Call Duration'].sum()

        # Add campaign header to the result1 string
        result1 += f"<h2>{campaign} (Campaign)</h2>"

        # Add HTML table for positive dispositions
        result1 += "<table border='1'>"
        result1 += "<tr><th>Positive Disposition</th><th>Count</th><th>% Age</th><th>Total Call Length</th></tr>"

        for disposition in positive_counts.index:
            result1 += f"<tr><td>{disposition}</td>"
            result1 += f"<td>{positive_counts[disposition]}</td>"
            result1 += f"<td>{positive_percentage[disposition]:.2f}%</td>"
            result1 += f"<td>{positive_total_call_length}</td></tr>"

        result1 += "</table>"

        # Add HTML table for negative dispositions
        result1 += "<table border='1'>"
        result1 += "<tr><th>Negative Disposition</th><th>Count</th><th>% Age</th><th>Total Call Length</th></tr>"

        for disposition in negative_counts.index:
            result1 += f"<tr><td>{disposition}</td>"
            result1 += f"<td>{negative_counts[disposition]}</td>"
            result1 += f"<td>{negative_percentage[disposition]:.2f}%</td>"
            result1 += f"<td>{negative_total_call_length}</td></tr>"

        result1 += "</table>"
    return render_template('index.html',  result1=result1, result2=result2, result_df=result_df)

if __name__ == '__main__':
    app.run(debug=True)
