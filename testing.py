import pandas as pd
import numpy as np
# Example DataFrame
df = pd.DataFrame({
    'A': [1, 2, 3],
    'B': ['x', 'y', 'z']
})

# Assign a floating-point value to a new column 'C' one row at a time
df['C'] = np.nan  # Initialize column 'C' with None values

# Assign floating-point values to 'C' one row at a time
df.loc[0, 'C'] = 1.5  # Assign float value to first row of column 'C'
df.loc[1, 'C'] = 2.7  # Assign float value to second row of column 'C'
df.loc[2, 'C'] = 3.9  # Assign float value to third row of column 'C'

# Display the updated DataFrame
print(df.info())
