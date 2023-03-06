import pandas as pd
import numpy as np

source_file = '/path/to/file'
destination_file = '/path/to/file'

float_digit = '%.4f'

# Read in data from source.xlsx file
df = pd.read_excel(source_file)

# Get the name of the first column
column1_name = df.columns[0]

# Interpolate data
step = 0.001
interpolated_data = pd.DataFrame({column1_name: np.arange(df[column1_name].min(), df[column1_name].max()+step, step)})
for column in df.columns[1:]:
    interpolated_data[column] = np.interp(interpolated_data[column1_name], df[column1_name], df[column])

# Write interpolated data to file
interpolated_data.to_excel(destination_file, index=False, float_format=float_digit, freeze_panes=(1, 1))
