from tkinter import filedialog, Tk
import pandas as pd

root = Tk()
root.withdraw()

csv_file = filedialog.askopenfilename(
    title = 'Select CSV file',
    filetypes = [("CSV Files", ".csv")]
)

if csv_file:
    df = pd.read_csv(csv_file)

    mask = df[['#', '#(1)', '#(2)']].notna().any(axis=1)
    
    # For those rows, concatenate Name with the # columns
    df.loc[mask, 'Name'] = (
        df.loc[mask, 'Name'].fillna('').astype(str) + ' ' +
        df.loc[mask, '#'].fillna('').astype(str) + ' ' +
        df.loc[mask, '#(1)'].fillna('').astype(str) + ' ' +
        df.loc[mask, '#(2)'].fillna('').astype(str)
    ).str.strip()
    
    # Save back to the same file (or change path for new file)
    df.to_csv(csv_file, index=False)
    print(f"Updated {csv_file}")