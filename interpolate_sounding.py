import pandas as pd

# Source and destination file paths
source_file = '/Path/to/file'
destination_file = '/Path/to/file'

# Float format
float_digit = '%.3f'

# Define the step size
step = 0.01

# Load the xls file into a pandas ExcelFile object
xls = pd.ExcelFile(source_file)

# Create a new ExcelWriter object for the destination file
with pd.ExcelWriter(destination_file) as writer:
    # Add an empty sheet named "CAL"
    pd.DataFrame().to_excel(writer, sheet_name='CAL')

    # Create a dataframe with the column names
    column_names_df = pd.DataFrame({
        'NAME': xls.sheet_names,
        'DRAFT': 0.00,
        'TRIM': 0.00,
        'VOLUME': ''
    })

    # Write the column names dataframe to the CAL sheet
    column_names_df.to_excel(writer, sheet_name='CAL', startrow=0, index=False, float_format=float_digit)

    # Iterate over all sheets in the source Excel file
    for sheet_name in xls.sheet_names:
        # Load the sheet into a pandas DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Set the index column to be the first column of the DataFrame
        df = df.set_index(df.columns[0])

        # Determine the starting and ending points based on the step size and the number of rows in the DataFrame
        start = df.index.min()
        end = df.index.max()
        num_rows = round((end - start) / step) + 1

        # Create a new DataFrame with the desired index values for interpolation
        new_index = pd.Index([start + i * step for i in range(num_rows)], name=df.index.name)
        new_df = pd.DataFrame(index=new_index)

        # Merge the original DataFrame with the new DataFrame and interpolate missing values
        interpolated_df = df.merge(new_df, how='outer', left_index=True, right_index=True).interpolate()

        # Transpose the DataFrame
        df_T = interpolated_df.T

        # Determine the starting and ending points based on the step size and the number of rows in the DataFrame
        start = df_T.index.min()
        end = df_T.index.max()
        num_rows = round((end - start) / step) + 1

        # Create a new DataFrame with the desired index values for interpolation
        new_index = pd.Index([start + i * step for i in range(num_rows)], name=df_T.index.name)
        new_df = pd.DataFrame(index=new_index)

        # Merge the original DataFrame with the new DataFrame and interpolate missing values
        interpolated_df_T = df_T.merge(new_df, how='outer', left_index=True, right_index=True).interpolate()

        # Transpose the result back to get the interpolated values for the rows
        interpolated_df = interpolated_df_T.T

        # Save the interpolated data to the corresponding sheet in the destination Excel file with the same sheet name
        interpolated_df.to_excel(writer, sheet_name=sheet_name, freeze_panes=(1, 1), float_format=float_digit)
