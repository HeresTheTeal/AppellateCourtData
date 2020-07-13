import os
import json
import csv
import sys
import openpyxl
from openpyxl.styles import Font
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import unidecode
from bs4 import BeautifulSoup
from datetime import datetime
import re


# TODO: Set directories, create time track variable for runtime analysis
# JSON Directory > (12 different folders) CA# > individual json files
# CSV Directory > (12 different files) ca##DataForSTM.csv
json_directory = 'C:\\Users\\Andrew\\Desktop\\Appellate Data\\Raw Data'
csv_directory = 'C:\\Users\\Andrew\\Desktop\\Appellate Data\\stmCSV'

# Create time variable to track runtime
startTime = datetime.now()


# TODO: Create spreadsheet to send output to
# Create spreadsheet
wb = openpyxl.Workbook()
ws = wb.active

# Name sheet, create header row
ws.title = 'Authoring Judge Data'
header_row = ['Index', 'File', 'Court', 'Judge List', 'Method', 'Authoring Judge', 'Opinion Text',
              'C / D Judge 1', 'C / D Type 1', 'C / D Text 1', 'C / D Judge 2', 'C / D Type 2', 'C / D Text 2',
              'C / D Judge 3', 'C / D Type 3', 'C / D Text 3']
ws.append(header_row)

# Bold header row
columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
           'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE']
for column in range(0, len(header_row)):
    cell = columns[column] + '1'
    cell_object = ws[cell]
    cell_object.font = Font(bold=True)

# Format column widths
ws.column_dimensions['A'].width = 5.5       # index
ws.column_dimensions['B'].width = 14        # file
ws.column_dimensions['C'].width = 7.5       # court
ws.column_dimensions['D'].width = 52.5      # judge list
ws.column_dimensions['E'].width = 15        # method
ws.column_dimensions['F'].width = 21        # authoring judge
ws.column_dimensions['G'].width = 45        # opinion text
ws.column_dimensions['H'].width = 21        # concur / dissent judge
ws.column_dimensions['I'].width = 14        # concur / dissent opinion type
ws.column_dimensions['J'].width = 45        # concur / dissent text
ws.column_dimensions['K'].width = 21        # concur / dissent judge
ws.column_dimensions['L'].width = 14        # concur / dissent opinion type
ws.column_dimensions['M'].width = 45        # concur / dissent text
ws.column_dimensions['N'].width = 21        # concur / dissent judge
ws.column_dimensions['O'].width = 14        # concur / dissent opinion type
ws.column_dimensions['P'].width = 45        # concur / dissent text

# For CSV reading later, expand max size
# This is the quick and dirty approach for PC
maxInt = sys.maxsize
while True:

    # decrease the maxInt value by factor 10
    # as long as the OverflowError occurs.
    try:
        csv.field_size_limit(maxInt)
        break
    except OverflowError:
        maxInt = int(maxInt/10)

# Set up CSV file
csv_file = open('Appellate Data ' + str(datetime.now().strftime("%d-%m-%Y %H-%M-%S")) + '.csv', 'w',
                newline='', encoding='utf-8')
csv_writer = csv.writer(csv_file)
csv_writer.writerow(header_row)


# TODO: Get text of each case
# Function to fetch the plain text from each json file
def get_html(directory, circuit_folder, file, html_field):

    # Load file
    path = os.path.join(directory, circuit_folder, file)
    with open(path, encoding='utf-8') as json_file:
        try:
            data = json.load(json_file)
        except json.decoder.JSONDecodeError:
            data = {html_field: 'ERROR - GET HTML'}

    # Get text
    html = data[html_field]
    return html


# TODO: Get panel judges from existing data
# Function to pull panel judges string from csv
def get_panel(csv_directory, csv_file, file):

    # Open csv
    with open(os.path.join(csv_directory, csv_file), encoding='utf-8') as csv_file_object:
        csv_reader = csv.DictReader(csv_file_object)

        # Find row with file in csv
        for row in csv_reader:
            if row['filename'] == file:

                # Create both string and list of judges
                judges_raw = row['judges']
                judges = judges_raw.split(', ')

                # Clean the list of judges to only use last name
                judges_list = []
                judge_names = []
                for judge in judges:
                    last, space, first = judge.partition(' ')
                    judges_list.append(last)
                    judge_names.append(last)
                    judge_names.append(first)

                # Return a list, where [0] is string, [1] is list
                return [judges_raw, judges_list, judge_names]

        # Error message if no match
        return 'ERROR - PANEL JUDGES'


# TODO: Find concur and dissent to be used in get_authoring_judge
# Create concur / dissent search method
def concur_dissent_search(lines_passed, line, judge_list, author_judge):

    # Empty variables to return unless something found
    lines_passed_new = lines_passed
    judge_output = ''
    type_output = ''
    match_output = ''

    # Create regex for date format of a concur / dissent we don't want in a proper concur / dissent line (ex: "1980)")
    dateRegex = re.compile(r'''(
                                       (\d{4})                                      # 4 digits for year
                                       ([)])                                        # close parentheses
                                       )''', re.VERBOSE)

    # First, check for of omitted words, avoiding false positive concur / dissents
    # Also check for concur or dissent up front to cut processing
    if lines_passed > 5 and (('concur' in line) or ('dissent' in line)) and ('filed' not in line) \
            and (' see ' not in line) and (' at ' not in line) and ('noting ' not in line) \
            and (' he ' not in line) and len(line) < 250 and (dateRegex.search(line) is None):

        # Loop each judge's name in each line
        for judge in judge_list:

            # Check if judge name + concur / dissent in a line to suggest concurrence / dissent
            if (judge in line) and ('dissent' in line) and ('concur' in line) and (judge != author_judge):
                lines_passed_new = 0
                judge_output = judge
                type_output = 'concur & dissent'
                match_output = line
                break

            if (judge in line) and ('dissent' in line) and (judge != author_judge):
                lines_passed_new = 0
                judge_output = judge
                type_output = 'dissent'
                match_output = line
                break

            if (judge in line) and ('concur' in line) and (judge != author_judge):
                lines_passed_new = 0
                judge_output = judge
                type_output = 'concur'
                match_output = line
                break

    # Return
    return [lines_passed_new, judge_output, type_output, match_output]


# TODO: Progress-guided search in get_authoring_judge
def progress_line_search(line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge,
                         concur_dissent_author_list, concur_dissent_type_list, match_line_list):

    # First, find line with all panel judges
    if progress == 'START':

        for judge in judge_list:
            if judge in line:
                judges_seen.append(judge)

        if all(judge in judges_seen for judge in judge_list):
            progress = 'AUTHOR'

        return (line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge,
                concur_dissent_author_list, concur_dissent_type_list, match_line_list)

    # Next, find line with any judge name, record author
    if progress == 'AUTHOR':

        # Author if per curiam
        if 'per curiam' in line:
            judge_found = True
            progress = 'CONCUR DISSENT'
            author_judge = 'per curiam'

        # If not per curiam
        else:

            # Loop for actual judges
            for judge in judge_list:

                # Below line checks for opinion author and skips if line announces concur / dissent
                if (judge in line) and ('dissent' not in line) and ('concur' not in line) and ('llp' not in line):
                    judge_found = True
                    progress = 'CONCUR DISSENT'
                    author_judge = judge
                    break

        return (line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge,
                concur_dissent_author_list, concur_dissent_type_list, match_line_list)

    # Now, search for any concurrence or dissent
    if progress == 'CONCUR DISSENT':

        # Lines passed
        lines_passed += 1

        # Call concur dissent search function
        concur_dissent_output = concur_dissent_search(lines_passed, line, judge_list, author_judge)
        lines_passed = concur_dissent_output[0]

        # Check if a concur / dissent found
        if concur_dissent_output[1] != '':
            concur_dissent_author_list.append(concur_dissent_output[1])
            concur_dissent_type_list.append(concur_dissent_output[2])
            match_line_list.append(concur_dissent_output[3])

        return (line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge,
                concur_dissent_author_list, concur_dissent_type_list, match_line_list)


# TODO: Get Authoring Judge
def get_authoring_judge(list, judge_list):

    # Output lists
    concur_dissent_author_list = []
    concur_dissent_type_list = []
    match_line_list = []

    # Author judge string to make sure concurrence / dissent author not same as authoring judge
    author_judge = ''

    # Variable to keep track if loop below has passed list of panel judges yet
    judge_found = False
    progress = 'START'

    # Loop the lines of the file until a judge is determined
    while judge_found is False:

        # Variable to count the lines since the main author found (to make sure early concur / dissent not picked up)
        lines_passed = 0

        # Variable to keep track of judges seen to address cases where judges on diff lines
        judges_seen = []

        # Uses two defined functions to find author, concur, dissent
        for line in list:

            # Variables to pass to function
            search_variables = [line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge,
                                concur_dissent_author_list, concur_dissent_type_list, match_line_list]

            # Call function and unpack into the same variables
            line, progress, judge_list, judges_seen, judge_found, lines_passed, author_judge, \
                concur_dissent_author_list, concur_dissent_type_list, \
                match_line_list = progress_line_search(*search_variables)

            # On a second loop (for per curiam default cases), skip the author step by advancing progress
            if author_judge == 'per curiam (default)':
                if progress == 'AUTHOR':
                    judge_found = True
                    progress = 'CONCUR DISSENT'

        # Return output
        if judge_found:
            output = [author_judge, concur_dissent_author_list, concur_dissent_type_list, match_line_list]
            return output

        # This is a second loop for per curiam (default) situations (where no judge spotted, assumed per curiam)
        else:
            author_judge = 'per curiam (default)'
            progress = 'START'


# TODO: Get Authoring Judge from HTML
def get_authoring_judge_html(html, judge_list):

    # Create new list
    new_text = BeautifulSoup(html, features='html.parser')
    new_text = new_text.get_text()
    text_list_raw = new_text.split('\n')

    # Clean text by making lower case and removing accents
    lower_text_list = []
    for string in text_list_raw:
        string = string.lower()
        string = unidecode.unidecode(string)
        lower_text_list.append(string)

    # Split newline characters in the HTML
    final_text_list = []
    for string in lower_text_list:

        # Replace duplicate newline characters
        string.replace('\n\n', '\n')

        # If the string in the list has a newline, split that string into a list by line
        if '\n' in string:
            sub_list = string.split('\n')

            # For each line in the string, append to the final list
            for sub_string in sub_list:
                final_text_list.append(sub_string)

        # If string doesn't have a new line (already a single line), append to final list
        else:
            final_text_list.append(string)

    return get_authoring_judge(final_text_list, judge_list)


# TODO: Get text from the csv for splitting for new file
def get_csv_text(csv_directory, csv_file, file):

    # Open csv
    with open(os.path.join(csv_directory, csv_file), encoding='utf-8') as csv_file_object:
        csv_reader = csv.DictReader(csv_file_object)

        # Find row with file in csv
        for row in csv_reader:
            if row['filename'] == file:

                # Create variable with string of csv text
                raw_csv_text = row['document']

                # Edit raw text
                raw_csv_text = raw_csv_text.replace('\n\n', '\n')
                csv_text_list = raw_csv_text.split('\n')

                # Return the list of csv text
                return [raw_csv_text, csv_text_list]

        # Error message if CSV text not found
        return 'ERROR - CSV TEXT'


# TODO: Split opinion and concur / dissent text
# noinspection DuplicatedCode
def split_text(match_line, csv_text_list, judge_names):

    # Create tracking variables
    split_tracker = 'SEARCH'
    opinion_type = ''

    # Create lists for initial reverse loop
    opinion_list = []
    concur_dissent_list = []

    # Create strings to put in list
    opinion_text = ''
    concur_dissent_text = ''

    # Create corrected match line by removing punctuation
    corrected_match_line = re.sub(r'[^\w\s]', '', match_line)

    # Determine if match line is concur or dissent
    if 'concur' in corrected_match_line:
        opinion_type = 'concur'
    elif 'dissent' in corrected_match_line:
        opinion_type = 'dissent'

    # Remove judge names from match line
    for name in judge_names:
        if name in corrected_match_line:
            corrected_match_line = corrected_match_line.replace(name, '')

    # Remove double spaces
    corrected_match_line = corrected_match_line.replace('  ', ' ')

    # Loop all lines in csv text IN REVERSE (to avoid finding a match too early)
    for line in reversed(csv_text_list):

        # If tracking variable is set to 'SPLIT,' add to opinion text string
        if split_tracker == 'SPLIT':
            opinion_list.append(line)
            continue

        # Corrected line by making lowercase
        corrected_line = line.lower()

        # Remove judge names from csv line
        for name in judge_names:
            if name in corrected_match_line:
                corrected_match_line = corrected_match_line.replace(name, '')

        # Try to cut out parts that won't match, or skip this line
        try:
            # Take off whitespace from either side
            corrected_line = corrected_line.strip()

            # Take out single characters
            while corrected_line[0].isalpha() and corrected_line[1].isspace():
                corrected_line = corrected_line[2:]

            # Take out punctuation
            corrected_line = re.sub(r'[^\w\s]', '', corrected_line)

        # If the corrected line is 2 char or less, skip
        except IndexError:
            continue

        # Remove double spaces
        corrected_line = corrected_line.replace('  ', ' ')

        # Take off whitespace once more
        corrected_line = corrected_line.strip()

        # If corrected line is 3 char or less, skip
        if len(corrected_line) < 4:
            concur_dissent_list.append(line)
            continue

        # Check if match line in current line
        if (corrected_line in corrected_match_line) and (opinion_type in corrected_line):

            # Set tracker, append text
            concur_dissent_list.append(line)
            split_tracker = 'SPLIT'
            continue

        # If not a concur / dissent
        else:
            concur_dissent_list.append(line)

    # Loop both lists and append to strings
    for line in reversed(opinion_list):
        opinion_text = opinion_text + line + '\n'
    for line in reversed(concur_dissent_list):
        concur_dissent_text = concur_dissent_text + line + '\n'

    # Error text if the concur / dissent text not picked up
    if opinion_text == '':
        opinion_text = concur_dissent_text
        concur_dissent_text = 'ERROR - CONCUR DISSENT TEXT NOT FOUND.'

    # Output list
    text_split_list = [opinion_text, concur_dissent_text, corrected_match_line]
    return text_split_list


# TODO: Actually process
# Create index for each row of spreadsheet
index = 1
total_files = 0

for circuit in os.listdir(json_directory):

    # Ignore hidden files in directory
    if 'CA' not in circuit:
        continue

    # Load in csv for circuit
    circuit_value = str(circuit.split('_')[1])
    circuit_folder = 'CA_' + circuit_value
    csv_file = 'ca' + circuit_value + 'DataForSTM.csv'
    with open(os.path.join(csv_directory, csv_file), encoding='utf-8') as csv_file_object:
        csv_reader = csv.DictReader(csv_file_object)

        # Create list of files in csv
        csv_json_files = []
        for row in csv_reader:
            csv_json_files.append(row['filename'])

        # Test message
        print('\n\n' + csv_file + '\n' + 'Total Files: ' + str(len(csv_json_files)) + '\n')
        progress_number = 0

        # Loop through files in each circuit
        for file in os.listdir(os.path.join(json_directory, circuit)):

            # Blank row to add data to if there is a match
            new_row = []

            # Check if file is in csv
            if file in csv_json_files:

                # Total progress counter
                total_files += 1

                # Print progress
                progress_number += 1
                if progress_number % 100 == 0:
                    print('*** ' + str(progress_number) + ' of ' + str(len(csv_json_files)) + ' ***')

                # Print file
                print(file)

                # If file in csv, add index to row
                new_row.append(str(index))
                index += 1

                # Add file, court to row
                new_row.append(file)
                new_row.append('CA ' + circuit_value)

                # Add panel judges to row using string from get_panel function
                panel = get_panel(csv_directory, csv_file, file)
                new_row.append(panel[0])

                # Get html, define method variable
                html = get_html(json_directory, circuit_folder, file, 'html')
                html_method = 'html'

                # Try different HTML fields in JSON if 'html' field empty
                if html == '':
                    html = get_html(json_directory, circuit_folder, file, 'html_lawbox')
                    html_method = 'html_lawbox'
                if html == '':
                    html = get_html(json_directory, circuit_folder, file, 'html_columbia')
                    html_method = 'html_columbia'
                if html == '':
                    html = get_html(json_directory, circuit_folder, file, 'html_with_citations')
                    html_method = 'html_with_citations'

                # Add method
                new_row.append(html_method)

                # Create html authoring judge variable, append main author
                html_judge_output = get_authoring_judge_html(html, panel[1])
                new_row.append(html_judge_output[0])

                # Define lists from html_judge_output to use in split text function
                concur_dissent_author_list = html_judge_output[1]
                concur_dissent_type_list = html_judge_output[2]
                match_line_list = html_judge_output[3]

                # First, check if there is no concurrence / dissent, and if not, append original case's text and blanks
                if not match_line_list:
                    new_row.append(get_csv_text(csv_directory, csv_file, file)[0])

                # If concurrence / dissent, process with split_text
                else:

                    # Get initial opinion text list
                    csv_text_list = get_csv_text(csv_directory, csv_file, file)[1]

                    # Create list to store all concur dissent data
                    concur_dissent_data = []

                    # Split text for each concurrence / dissent, going through match values in reverse
                    for match in range(len(match_line_list) - 1, -1, -1):

                        # First, run split_text, starting with the last match and moving forward
                        text_split_list = split_text(match_line_list[match], csv_text_list, panel[2])

                        # Check if this is the final match
                        if match == 0:

                            # Add to concur dissent data list
                            concur_dissent_data.append(text_split_list[1])
                            concur_dissent_data.append(concur_dissent_type_list[match])
                            concur_dissent_data.append(concur_dissent_author_list[match])

                            # Append regular opinion text
                            new_row.append(text_split_list[0])

                            # Append all concur / dissent data
                            for string in reversed(concur_dissent_data):
                                new_row.append(string)

                        # If there are still additional matches to find
                        else:

                            # Redeclare csv_text_list for subsequent search
                            csv_text_list = text_split_list[0].split('\n')

                            # Append concur / dissent to data list
                            concur_dissent_data.append(text_split_list[1])
                            concur_dissent_data.append(concur_dissent_type_list[match])
                            concur_dissent_data.append(concur_dissent_author_list[match])

                # Extend new row list with empty values until the proper length
                while len(new_row) < 16:
                    new_row.append('')

                # Remove illegal characters from new_row for .xlsx output (since illegal characters throw error)
                clean_new_row = []
                for string in new_row:

                    # New string variable
                    new_string = string

                    # Check for illegal characters and replace
                    if ILLEGAL_CHARACTERS_RE.search(string) is not None:
                        new_string = re.sub(ILLEGAL_CHARACTERS_RE, '', string)

                    # Create clean new row
                    clean_new_row.append(new_string)

                # Add new row to spreadsheets
                ws.append(clean_new_row)
                csv_writer.writerow(new_row)


# TODO: Save and exit
# Save
wb.save('Appellate PLUS TEXT - ' + str(datetime.now().strftime("%m-%d-%Y %H-%M-%S")) + '.xlsx')

# Final message
print('* Complete. File saved. *')

# Runtime and per-unit
runtime = datetime.now() - startTime
total_seconds = runtime.total_seconds()
seconds_per_file = total_seconds / total_files
print('Runtime: ' + str(runtime))
print('Runtime, seconds per file: ' + str(seconds_per_file))
