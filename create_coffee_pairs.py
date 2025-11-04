import pandas as pd
import random
import os
import sys
import glob
import re

# --- 1. Configuration & File Names ---
SHEET_NAME = "All Exe & Ass" # The sheet with the member list
PAST_MATCHES_PREFIX = "past_matches_"
PAST_MATCHES_SUFFIX = ".csv"
CURRENT_MATCHES_FILE = "current_matches.txt"

# --- 2. Helper Functions ---

def get_member_file():
    """
    Gets the Excel file path.
    1. Checks for a command-line argument.
    2. If none, searches for the first .xlsx file in the script's directory.
    """
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()

    if len(sys.argv) > 1:
        filename = sys.argv[1]
        print(f"Using provided filename: {filename}")
        if not os.path.exists(filename):
            print(f"Error: File not found at specified path: {filename}")
            return None
        return filename
    
    else:
        print("No filename provided, searching for *.xlsx in script directory...")
        search_path = os.path.join(script_dir, '*.xlsx')
        excel_files = glob.glob(search_path)
        
        if excel_files:
            filename = excel_files[0] # Use the first one found
            print(f"Found and using file: {filename}")
            return filename
        else:
            print(f"Error: No .xlsx file found in directory: {script_dir}")
            return None

def load_member_data(member_file):
    """
    Loads and cleans the current member list from the given Excel file.
    If duplicates or missing columns are found, prints an error and exits.
    """
    try:
        df = pd.read_excel(
            member_file, 
            sheet_name=SHEET_NAME, 
            header=0 
        )
        
        # --- 1. CLEAN COLUMN HEADERS ---
        # This turns "Exec or Assoc " into "Exec or Assoc"
        df.columns = df.columns.str.strip()
        
        # --- 2. CHECK FOR REQUIRED COLUMNS ---
        REQUIRED_COLUMNS = ['Full Name', 'Department', 'Exec or Assoc']
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        
        if missing_cols:
            print("----------------------------------------------------------------")
            print(f"❌ FATAL ERROR: Missing required columns in '{member_file}':")
            print(f"  Missing: {', '.join(missing_cols)}")
            print(f"  Available columns are: {list(df.columns)}")
            print("\nPlease ensure the 'All Exe & Ass' sheet has 'Full Name', 'Department', and 'Exec or Assoc' columns.")
            print("----------------------------------------------------------------")
            sys.exit(1)

        # --- 3. CLEAN DATA (TRIM WHITESPACE) ---
        # First, drop any rows that are completely empty in 'Full Name'
        df = df.dropna(subset=['Full Name'])
        
        # Now, trim whitespace from all required string columns
        for col in REQUIRED_COLUMNS:
            # .astype(str) handles any non-string data (like numbers)
            # .str.strip() removes whitespace from start and end
            df[col] = df[col].astype(str).str.strip()

        # --- 4. CHECK FOR DUPLICATES (on cleaned data) ---
        duplicates = df[df.duplicated(subset=['Full Name'], keep=False)]
        
        if not duplicates.empty:
            duplicate_names = duplicates['Full Name'].unique()
            print("----------------------------------------------------------------")
            print(f"❌ FATAL ERROR: Duplicate names found in '{member_file}':")
            for name in duplicate_names:
                print(f"  - {name}")
            print("\nPlease correct the Excel file before running again.")
            print("----------------------------------------------------------------")
            sys.exit(1)

        # --- 5. CONVERT TO DICTIONARY ---
        # This code now only runs if no duplicates are found
        member_lookup = df.to_dict('index')
        all_people_ids = list(member_lookup.keys())
        
        return all_people_ids, member_lookup

    except FileNotFoundError:
        print(f"❌ FATAL ERROR: Could not find member file: {member_file}") 
        sys.exit(1)
    except Exception as e:
        print(f"❌ FATAL ERROR: Error reading Excel file: {e}")
        print("This can happen if the sheet name 'All Exe & Ass' is incorrect or the file is corrupt.")
        sys.exit(1)


def load_past_matches(): # <-- MODIFIED
    """
    Loads and cleans all pairs from ALL versioned history files.
    Returns a set of all past pairs and the next version number to use.
    """
    past_pairs = set()
    max_version = 0
    
    pattern = re.compile(f"{PAST_MATCHES_PREFIX}(\\d+){PAST_MATCHES_SUFFIX}")
    
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
        
    search_path = os.path.join(script_dir, f"{PAST_MATCHES_PREFIX}*{PAST_MATCHES_SUFFIX}")
    history_files = glob.glob(search_path)
    
    if history_files:
        print(f"Found {len(history_files)} history file(s)...")
    else:
        print("No history files found. Starting new history.")

    for file_path in history_files:
        match = pattern.search(os.path.basename(file_path))
        
        if not match:
            print(f"Skipping file with unexpected name: {file_path}")
            continue

        try:
            version_num = int(match.group(1))
            if version_num > max_version:
                max_version = version_num
            
            df_past = pd.read_csv(file_path)
            
            # --- CLEAN DATA FROM CSV ---
            if 'Person1' not in df_past.columns or 'Person2' not in df_past.columns:
                print(f"Warning: Skipping history file {file_path}, missing 'Person1' or 'Person2' column.")
                continue
                
            # Trim whitespace just in case CSV was manually edited
            df_past['Person1'] = df_past['Person1'].astype(str).str.strip()
            df_past['Person2'] = df_past['Person2'].astype(str).str.strip()
            # --- END CLEAN ---

            for _, row in df_past.iterrows():
                # Now using cleaned names
                pair = frozenset([row['Person1'], row['Person2']])
                past_pairs.add(pair)
                
        except Exception as e:
            print(f"Error reading or parsing history file {file_path}: {e}")
    
    next_version = max_version + 1
    print(f"Total {len(past_pairs)} unique past pairs loaded.")
    print(f"Next file version will be: {next_version} ({PAST_MATCHES_PREFIX}{next_version}{PAST_MATCHES_SUFFIX})")
    
    return past_pairs, next_version


def is_valid_pair(p1_id, p2_id, member_lookup, past_pairs):
    """
    Checks if a pair (by their IDs) is valid based on all your rules.
    All data from member_lookup is now pre-cleaned.
    """
    p1_data = member_lookup[p1_id]
    p2_data = member_lookup[p2_id]
    
    p1_name = p1_data['Full Name'] # This is now trimmed
    p2_name = p2_data['Full Name'] # This is now trimmed
    
    # 1. Check for past matches (compares trimmed name to trimmed history)
    if frozenset([p1_name, p2_name]) in past_pairs:
        return False
        
    # 2. Check for same department (using trimmed data)
    if p1_data['Department'] == p2_data['Department']:
        return False

    # 3. Check for Exec + Exec (using trimmed data)
    # This will now correctly find 'Exec or Assoc' column
    if p1_data['Exec or Assoc'] == 'Executive' and p2_data['Exec or Assoc'] == 'Executive':
        return False

    return True

# --- 3. Main Pairing Logic ---

def create_new_pairs():
    member_file = get_member_file()
    if not member_file:
        print("Exiting script. Please provide an Excel file.")
        return

    people_to_pair, member_lookup = load_member_data(member_file)
    past_pairs, next_version = load_past_matches()
    
    random.shuffle(people_to_pair)
    new_pairs = [] # Stores tuples of (p1_id, p2_id)
    unmatched = [] # Stores unmatched p_ids

    print(f"Starting with {len(people_to_pair)} people...")

    while len(people_to_pair) > 1:
        p1 = people_to_pair.pop(0)
        found_match = False
        
        for i, p2 in enumerate(people_to_pair):
            if is_valid_pair(p1, p2, member_lookup, past_pairs):
                new_pairs.append((p1, p2))
                people_to_pair.pop(i) 
                found_match = True
                break 
        
        if not found_match:
            unmatched.append(p1)

    unmatched.extend(people_to_pair)

    # --- 4. Output and Save Results ---
    print("\n--- New Coffee Chat Pairs ---")
    
    # These lists will hold the *names* for saving
    pairs_for_csv_history = []
    lines_for_txt_output = []

    for (p1_id, p2_id) in new_pairs:
        p1_name = member_lookup[p1_id]['Full Name']
        p2_name = member_lookup[p2_id]['Full Name']
        
        # 1. Print to console
        print(f"{p1_name}  <-->  {p2_name}")
        
        # 2. Prepare for saving
        pairs_for_csv_history.append((p1_name, p2_name))
        lines_for_txt_output.append(f"{p1_name} <--> {p2_name}")

    print("\n--- Unmatched People (Odd one out) ---")
    unmatched_names = []
    for p_id in unmatched:
        p_name = member_lookup[p_id]['Full Name']
        print(p_name)
        unmatched_names.append(p_name)
            
    # --- 5. Save results to files ---
    try:
        # --- 5a. Save to VERSIONED HISTORY (CSV) ---
        new_pairs_df = pd.DataFrame(pairs_for_csv_history, columns=['Person1', 'Person2'])
        new_file_name = f"{PAST_MATCHES_PREFIX}{next_version}{PAST_MATCHES_SUFFIX}"
        
        new_pairs_df.to_csv(
            new_file_name,
            index=False 
        )
        print(f"\nSuccessfully saved {len(new_pairs)} new pairs to history file: {new_file_name}")
        
        # --- 5b. Save to CURRENT MATCHES (TXT) ---
        # This is the new, human-readable file that overwrites itself
        with open(CURRENT_MATCHES_FILE, 'w', encoding='utf-8') as f:
            f.write("--- New Coffee Chat Pairs ---\n")
            for line in lines_for_txt_output:
                f.write(f"{line}\n")
            
            if unmatched_names:
                f.write("\n--- Unmatched People (Odd one out) ---\n")
                for name in unmatched_names:
                    f.write(f"{name}\n")
        
        print(f"Successfully saved current matches to readable file: {CURRENT_MATCHES_FILE}")
        
    except Exception as e:
        print(f"\nError saving new pairs: {e}")

# --- Run the script ---
if __name__ == "__main__":
    create_new_pairs()

