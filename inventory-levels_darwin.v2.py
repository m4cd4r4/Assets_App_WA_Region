import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import argparse

# Argument parsing for output file path
parser = argparse.ArgumentParser(description="Generate inventory level plot.")
parser.add_argument("--output", required=True, help="Path to save the output plot")
args = parser.parse_args()

# Construct the path to the file
file_path = os.path.join(os.path.dirname(__file__), 'EUC_Perth_Assets.xlsx')

# Load the spreadsheet
try:
    xl = pd.ExcelFile(file_path)
    df_items = xl.parse('Darwin_Items')
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please ensure the file exists.")
    sys.exit(1)
except Exception as e:
    print(f"Error loading data: {e}")
    sys.exit(1)

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
try:
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
    plt.title(f'Darwin - Inventory Levels - {current_date}', fontsize=14)
    plt.tight_layout()

    # Save the plot to the specified output file
    output_path = args.output
    plt.savefig(output_path)
    print(f"Plot saved at {output_path}")

    # Show the plot interactively
    plt.show()  # This will display the plot in a GUI window

except Exception as e:
    print(f"Error generating chart: {e}")
finally:
    # Ensure all plots are closed properly
    plt.close('all')
    print("Plot closed successfully.")
