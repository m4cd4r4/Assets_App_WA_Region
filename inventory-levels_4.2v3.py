import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import argparse

# Argument parsing for output file path
parser = argparse.ArgumentParser(description="Generate inventory level plot for Build Room.")
parser.add_argument("--output", required=True, help="Path to save the output plot")
args = parser.parse_args()

# Check if the application is "frozen"
if getattr(sys, 'frozen', False):
    # If it's frozen, use the path relative to the executable
    application_path = sys._MEIPASS
else:
    # If it's not frozen, use the path relative to the script file
    application_path = os.path.dirname(__file__)

# Construct the path to the file
file_path = os.path.join(application_path, 'EUC_Perth_Assets.xlsx')

# Load the spreadsheet
try:
    print(f"Loading spreadsheet from {file_path}")
    xl = pd.ExcelFile(file_path)
    df_items = xl.parse('BR_Items')
    print("Spreadsheet loaded successfully.")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please ensure the file exists.")
    sys.exit(1)
except Exception as e:
    print(f"Error loading data: {e}")
    sys.exit(1)

# Replace NaN values with 0 in 'NewCount' column
try:
    if 'NewCount' in df_items.columns:
        df_items['NewCount'] = df_items['NewCount'].fillna(0)
        print("NaN values in 'NewCount' column replaced with 0.")
    else:
        raise KeyError("'NewCount' column not found in the dataset.")
except KeyError as e:
    print(f"Error processing 'NewCount' column: {e}")
    sys.exit(1)

# Ensure 'Item' column exists in the dataset
if 'Item' not in df_items.columns:
    print("Error: 'Item' column not found in the dataset.")
    sys.exit(1)

# Create a horizontal bar chart for the current inventory levels
try:
    print("Generating the plot...")
    plt.figure(figsize=(14 * 0.60, 10 * 0.60))
    bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

    # Add the text with the count at the end of each bar
    for bar in bars:
        width = bar.get_width()
        width_int = int(width)
        plt.text(width + 1, bar.get_y() + bar.get_height() / 2, width_int, ha='left', va='center', color='black')

    plt.ylabel('Item', fontsize=12)
    plt.xlabel('Volume', fontsize=12)
    plt.xlim(0, df_items['NewCount'].max() + 20)  # Dynamically adjust x-axis limit
    current_date = datetime.now().strftime('%d-%m-%Y')
    plt.title(f'Build Room - Inventory Levels (Perth) - {current_date}', fontsize=14)
    plt.tight_layout()

    # Save the plot to the specified output file
    output_path = os.path.abspath(args.output)
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        print(f"Output directory {output_dir} does not exist. Creating it.")
        os.makedirs(output_dir, exist_ok=True)

    plt.savefig(output_path)
    print(f"Plot saved successfully at {output_path}")

except Exception as e:
    print(f"Error generating chart: {e}")
finally:
    # Ensure all plots are closed properly
    plt.close('all')
    print("Plot closed successfully.")

# End the script gracefully
print("Exiting application.")
