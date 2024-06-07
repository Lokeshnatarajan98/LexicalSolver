# Synopsis: This Python code is designed to load data from Excel files and search for a specific word in different categories (Noun, Adjective, Adverb, Verb). It performs the following steps:

# 1. It imports the pandas library as `pd` for data manipulation and the time library for time-related functions.
# 2. The `load_and_search_word` function is defined to load data from an Excel file, clean and preprocess the data, and then search for a target word within that data.
# 3. The `check_list` contains a list of words and values that should be ignored during the search.
# 4. The user is prompted to enter a jumbled word, which is then converted to uppercase and sorted alphabetically. This sorted version is stored in the variable `y`.
# 5. Excel file paths are stored in the `files` dictionary, categorized by word type.
# 6. A loop iterates through the `files` dictionary, and for each file, it calls the `load_and_search_word` function, passing the file path, the word type, and the sorted jumbled word. The results are stored in the `results` dictionary.
# 7. The `print_result` function is defined to print the results. It checks if there is at least one row in the DataFrame (`df`) and, if so, prints the found word and its associated word type. If no results are found, it prints a message indicating that the word was not found in that category.
# 8. Finally, the script records and prints the start and end times for the entire script.

import pandas as pd
import time


# Function to load and search word
def load_and_search_word(file_path, word_type, target_word):
    curr_time = time.strftime("%H:%M:%S", time.localtime())
    print(f'Start time ({word_type}): {curr_time}')
    
    df = pd.read_excel(file_path)
    df['dictword'] = df['dictword'].str.strip().str.upper()
    df = df[~df['dictword'].str.upper().isin(check_list)]
    df['reversed'] = df['dictword'].apply(lambda x: ''.join(sorted(str(x))))
    found_word_df = df[df['reversed'].str.strip().str.upper() == target_word]
    
    curr_time = time.strftime("%H:%M:%S", time.localtime())
    #print(f'Start time (Script): {script_start_time}')
    print(f'End time ({word_type}): {curr_time}')
    
    return found_word_df

# Record start time for the entire script
script_start_time = time.strftime("%H:%M:%S", time.localtime())

check_list = ['NONE', 'NULL', 'NAN', 'NAT', "", None, " "]
y = ''.join(sorted(str(input('Enter the jumble puzzle word: ')).upper()))

files = {
    'Noun': r'C:\python\Project1\Nounfile1.csv.xlsx',
    'Adjective': r'C:\python\Project1\adjfile.csv.xlsx',
    'Adverb': r'C:\python\Project1\advfile.csv.xlsx',
    'Verb': r'C:\python\Project1\verbfile.csv.xlsx',
}

results = {}
for word_type, file_path in files.items():
    results[word_type] = load_and_search_word(file_path, word_type, y)

def print_result(word_type, df):
    if df.shape[0] > 0:
        print(f'{df["dictword"].values[0]} {word_type}')
    else:
        print(f'Not found in {word_type} file')

for word_type, df in results.items():
    print_result(word_type, df)

# Record end time for the entire script
script_end_time = time.strftime("%H:%M:%S", time.localtime())
print(f'Start time (Script): {script_start_time}')
print(f'End time (Script): {script_end_time}')

# Here's the flow of the code:
# - It starts by prompting the user to enter a jumbled word.
# - It then reads data from multiple Excel files, cleans the data, and searches for the jumbled word within each file for different word types (Noun, Adjective, Adverb, Verb).
# - The results are printed for each word type, indicating whether the word was found or not.
# - The script also records and prints the start and end times to measure the script's execution duration.
# The primary purpose of this code is unjumbling words by sorting the letters and searching for them in different word categories within Excel files. It's a useful tool for word puzzles or word-related applications.


