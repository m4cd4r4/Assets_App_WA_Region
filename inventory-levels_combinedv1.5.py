import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import argparse
from config import workbook_path  # Import the path from the config file

# Argument parsing for output file path
parser = argparse.ArgumentParser(description="Generate combined inventory level plot.")
parser.add_argument("--output", required=True, help="Path to save the output plot")
args = parser.parse_args()

# Use the imported workbook_path directly
file_path = workbook_path

# Load the spreadsheet
try:
    print(f"Loading spreadsheet from {file_path}")
    xl = pd.ExcelFile(file_path)
    df_42_items = xl.parse('4.2_Items')
    df_br_items = xl.parse('BR_Items')
    df_darwin_items = xl.parse('Darwin_Items')  # Load the Darwin_Items sheet
    print("Spreadsheet loaded successfully.")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please ensure the file exists.")
    exit(1)
except Exception as e:
    print(f"Error loading data: {e}")
    exit(1)

# Replace NaN values with 0 in 'NewCount' column for all dataframes
try:
    df_42_items['NewCount'] = df_42_items['NewCount'].fillna(0)
    df_br_items['NewCount'] = df_br_items['NewCount'].fillna(0)
    df_darwin_items['NewCount'] = df_darwin_items['NewCount'].fillna(0)  # Handle NaN for Darwin_Items
    print("NaN values in 'NewCount' columns replaced with 0.")
except Exception as e:
    print(f"Error processing 'NewCount' columns: {e}")
    exit(1)

# Combine the three dataframes
try:
    combined_df = pd.concat([
        df_42_items[['Item', 'NewCount']],
        df_br_items[['Item', 'NewCount']],
        df_darwin_items[['Item', 'NewCount']]  # Include Darwin_Items
    ])
    grouped_df = combined_df.groupby('Item', as_index=False)['NewCount'].sum()
    print("Data combined and grouped successfully.")
except KeyError as e:
    print(f"Error combining data: {e}")
    exit(1)

# Create a horizontal bar chart for the summed inventory levels
try:
    print("Generating the plot...")
    plt.figure(figsize=(14 * 0.60, 10 * 0.60))
    bars = plt.barh(grouped_df['Item'], grouped_df['NewCount'], color='#006aff')

    # Define the spacing for the text
    spacing = 1  # Adjust this value for more or less spacing

    # Add the text with the summed count at the end of each bar
    for bar in bars:
        width = bar.get_width()
        plt.text(width + spacing, bar.get_y() + bar.get_height() / 2,
                 f'{int(width)}', ha='left', va='center', color='black')

    plt.ylabel('Item', fontsize=12)
    plt.xlabel('Volume', fontsize=12)
    plt.xlim(0, grouped_df['NewCount'].max() + 20)  # Adjust dynamically based on data
    current_date = datetime.now().strftime('%d-%m-%Y')
    plt.title(f'Combined: B4.2, Build Room & Darwin - {current_date}', fontsize=14)
    plt.tight_layout()

    # Ensure 'Plots' folder exists
    output_dir = os.path.dirname(args.output)
    if not os.path.exists(output_dir):
        print(f"Output directory {output_dir} does not exist. Creating it.")
        os.makedirs(output_dir, exist_ok=True)

    # Save the plot to the specified output file
    plt.savefig(args.output)
    print(f"Plot saved successfully at {args.output}")

except Exception as e:
    print(f"Error generating chart: {e}")
finally:
    # Ensure all plots are closed properly
    plt.close('all')
    print("Plot closed successfully.")

# End the script gracefully
print("Exiting application.")
