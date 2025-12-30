import pandas as pd
import matplotlib.pyplot as plt

def generate_report(input_file, output_file):
    # Read CSV
    df = pd.read_csv(input_file)

    # Basic cleaning
    df = df.dropna()

    # Example: assume numeric column named 'value'
    if 'value' not in df.columns:
        raise ValueError("CSV must contain a 'value' column")

    summary = df['value'].describe()

    # Save to Excel
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        summary.to_frame(name='Summary').to_excel(writer, sheet_name='Summary')

        workbook = writer.book
        worksheet = writer.sheets['Summary']

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Values',
            'categories': ['Data', 1, 0, len(df), 0],
            'values':     ['Data', 1, df.columns.get_loc('value'), len(df), df.columns.get_loc('value')],
        })

        worksheet.insert_chart('D2', chart)

    print(f"Report saved to {output_file}")


if __name__ == "__main__":
    generate_report("sample_input.csv", "output_report.xlsx")
