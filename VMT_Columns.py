import pandas as pd
import os
import re

def read_file(file_path):
    """
    Reads a file based on its extension (csv or xlsx)

    Args:
        file_path (str): Path to the file to read
    Returns:
        pd.DataFrame: DataFrame containing the file data
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        return pd.read_csv(file_path)
    elif file_extension == '.xlsx':
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .csv or .xlsx files.")

def save_file(df, file_path):
    """
    Saves a DataFrame to a file based on its extension (csv or xlsx)

    Args:
        df (pd.DataFrame): DataFrame to save
        file_path (str): Path where to save the file
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        df.to_csv(file_path, index=False)
    elif file_extension == '.xlsx':
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .csv or .xlsx files.")

def remove_column_numbering(df):
    """
    Removes trailing decimal numbers from column names while preserving the data.
    Example: Converts columns like:
    'Custom field (Using Legal Entity (Application)).1',
    'Custom field (Using Legal Entity (Application)).2'
    to:
    'Custom field (Using Legal Entity (Application))',
    'Custom field (Using Legal Entity (Application))'

    Args:
        df (pd.DataFrame): DataFrame with numbered columns
    Returns:
        pd.DataFrame: DataFrame with cleaned column names
    """
    # Create a mapping of old column names to new ones (without the trailing numbers)
    column_mapping = {}
    for col in df.columns:
        # Remove trailing .number pattern
        new_col = re.sub(r'\.\d+$', '', col)
        column_mapping[col] = new_col

    # Rename the columns while preserving the data
    df.columns = [column_mapping[col] for col in df.columns]
    return df

def standardize_data(base_file, new_file):
    """
    Standardizes the structure of a new data file to match the base file.
    Handles repeated column headers (instances) and preserves their data.
    Only adds number suffixes to columns with multiple instances.
    Adds missing columns with empty values.

    Args:
        base_file (str): Path to the base file (csv or xlsx)
        new_file (str): Path to the new file to standardize (csv or xlsx)
    """
    try:
        # Read the base file to get the reference structure
        base_df = read_file(base_file)
        base_columns = base_df.columns.tolist()

        # Count instances of each column in base file
        base_column_counts = {}
        for col in base_columns:
            base_column_counts[col] = base_column_counts.get(col, 0) + 1

        # Read the new file
        new_df = read_file(new_file)
        new_columns = new_df.columns.tolist()

        # Count instances of each column in new file
        new_column_counts = {}
        for col in new_columns:
            new_column_counts[col] = new_column_counts.get(col, 0) + 1

        # Check if any column has fewer instances in new file
        missing_instances = []
        for col, count in base_column_counts.items():
            if col not in new_column_counts or new_column_counts[col] < count:
                missing_instances.append((col, count, new_column_counts.get(col, 0)))

        if missing_instances:
            print("WARNING: The following columns have insufficient instances in the new file:")
            for col, required, actual in missing_instances:
                print(f"- {col}: Required {required} instances, found {actual} instances")
            print("\nEmpty columns will be added for missing instances.")

        # Create standardized DataFrame
        standardized_df = pd.DataFrame()

        # Process each unique column (removing duplicates from base_columns)
        unique_columns = []
        processed_columns = set()

        for col in base_columns:
            if col not in processed_columns:
                unique_columns.append(col)
                processed_columns.add(col)

        # Process each column and its instances
        for col in unique_columns:
            # Get all instances of this column from new file
            new_col_instances = [c for c in new_columns if c == col]
            required_count = base_column_counts[col]

            # If column has only one instance, don't add number suffix
            if required_count == 1:
                if new_col_instances:
                    standardized_df[col] = new_df[new_col_instances[0]]
                else:
                    # Add empty column if missing
                    standardized_df[col] = pd.Series(dtype='object')
            else:
                # Add each instance with number suffix
                for i in range(required_count):
                    if i < len(new_col_instances):
                        standardized_df[f"{col}_{i+1}"] = new_df[new_col_instances[i]]
                    else:
                        # Add empty column for missing instances
                        standardized_df[f"{col}_{i+1}"] = pd.Series(dtype='object')

        # Remove trailing numbers from column names before saving
        standardized_df = remove_column_numbering(standardized_df)

        try:
            # Save the standardized data back to the original file
            save_file(standardized_df, new_file)
            print(f"File {new_file} has been standardized successfully!")
            print("Column instances preserved as per original data.")
            print("Missing columns added with empty values.")
            print("Trailing decimal numbers removed from column names.")
        except PermissionError:
            print(f"\nERROR: Could not save the file '{new_file}'.")
            print("Please make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")
            return

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        if "Permission denied" in str(e):
            print("\nPlease make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")

if __name__ == "__main__":
    standardize_data("Power BI Export VM (Vulnerability Management Tool & Risk Register) 2024-07-09T10_36_58+0200.xlsx", "Power BI Export VM (Vulnerability Management Tool & Risk Register) 2025-03-03T11_43_25+0100.xlsx")
