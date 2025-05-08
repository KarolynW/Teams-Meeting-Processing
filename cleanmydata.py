import os
import pandas as pd

# Set this to your local folder with the CSVs
FOLDER_PATH = 'Attendance Lists'
OUTPUT_FILE = 'combined_attendees.xlsx'


def extract_participants_section(file_path):
    with open(file_path, 'r', encoding='utf-16') as f:
        lines = f.readlines()

    section_2_data = []
    in_participants = False

    for line in lines:
        line = line.strip()
        if line.startswith('2.'):
            in_participants = True
            continue
        elif line.startswith('3.'):
            break  # End of section 2

        if in_participants:
            if line:  # skip empty lines
                section_2_data.append(line.split('\t'))

    if section_2_data:
        header = section_2_data[0]
        rows = section_2_data[1:]
        df = pd.DataFrame(rows, columns=header)
        return df
    return pd.DataFrame()

def main():
    all_data = []

    for file_name in os.listdir(FOLDER_PATH):
        if file_name.endswith('.csv'):
            full_path = os.path.join(FOLDER_PATH, file_name)
            df = extract_participants_section(full_path)
            if not df.empty:
                df['Source File'] = file_name
                all_data.append(df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df.to_excel(OUTPUT_FILE, index=False)
        print(f"Saved combined attendance to: {OUTPUT_FILE}")
    else:
        print("No attendance data found.")

if __name__ == "__main__":
    main()
