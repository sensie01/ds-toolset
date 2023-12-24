import sys
import pandas as pd
import warnings

def calculate_total_working_time(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path, header=None, skiprows=5)

        # Check if the specified columns are present
        if len(df.columns) < 3:
            raise ValueError("Columns 'Last Occurred' and 'Cleared On' not found. Ensure the data starts from row 6.")

        # Extract the relevant columns
        last_occurred_column = df.iloc[:, 1]
        cleared_on_column = df.iloc[:, 2]

        # Suppress the datetime format inference warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Convert datetime strings to datetime objects
            last_occurred_column = pd.to_datetime(last_occurred_column, errors='coerce')
            cleared_on_column = pd.to_datetime(cleared_on_column, errors='coerce')

        # Calculate the working period
        working_period = cleared_on_column - last_occurred_column

        # Sum up all individual working periods
        total_working_time = working_period.sum()

        # Print the total working time
        print(f"{total_working_time}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Check if a file path is provided as a command-line argument
    if len(sys.argv) != 2:
        print("Please run the program in this way: python calculate.py <the_file_path>")
    else:
        # Extract the file path from the command-line argument
        file_path = sys.argv[1]

        # Call the function with the provided file path
        calculate_total_working_time(file_path)
