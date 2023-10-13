"""
Module for processing university schedules from Excel files.

Classes:
    Schedule: Class for processing and storing university schedules.

Methods:
    parse_arguments: Function to parse command-line arguments for the script.

Usage:
    python script_name.py file1.xlsx file2.xlsx -o output_file.json
"""
import argparse
import json
import re
import pandas as pd


class Schedule:
    """
        Class for processing and storing university schedules.

        Attributes:
            DAYS_NAME (list): List of day names.
            _schedule_dict (dict): Dictionary to store the processed schedule.

        Methods:
            __init__: Initializes the Schedule class.
            _extract_specialties_and_year: Extracts specialties and study year from a text.
            _is_valid_time_range: Checks if the input string is a valid time range.
            _get_specs_in_str: Extracts specialties mentioned in a string and returns the cleaned string.
            _add_space_after_parenthesis: Adds a space after closing parenthesis if not present.
            _get_groups: Extracts groups from the input string.
            _get_week_array: Processes the input string to get an array of weeks.
            _proccess_time_range: Processes the input string to get a dictionary of start, end, and full time.
            _proccess_room: Processes the room information.
            _add_to_table: Adds data to the schedule table.
            procces_excels: Processes Excel files and populates the schedule dictionary.
            save_to_json: Saves the schedule dictionary to a JSON file.
            get_schedule_dict: Returns the schedule dictionary.
    """
    DAYS_NAME = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', 'П`ятниця', 'Субота', 'Неділя']
    _schedule_dict = {}

    def __init__(self, files: list | str):
        if isinstance(files, str):
            files: list = [files]
        self.procces_excels(files)

    @staticmethod
    def _extract_specialties_and_year(text):
        """
        Extracts specialties and study year from a text.

        Args:
            text (str): Input text.

        Returns:
            tuple: Tuple containing a list of specialties and the study year.
        """
        specialty_matches = re.findall(r'[«"](.*?)[»"]', text)
        year_matches = re.findall(r'(\d+) р\.н\.', text)

        specialties = [spec.strip() for spec in specialty_matches]
        years = int(year_matches[0])

        return specialties, years

    @staticmethod
    def _is_valid_time_range(input_str: str):
        """
        Checks if the input string is a valid time range.

        Args:
            input_str (str): Input time range string.

        Returns:
            bool: True if the input is a valid time range, False otherwise.
        """
        pattern = re.compile(r'^\d{1,2}[.:]\d{2}-\d{1,2}[.:]\d{2}$')
        return bool(pattern.match(input_str))

    @staticmethod
    def _get_specs_in_str(input_string, specs):
        """
        Extracts specialties mentioned in a string and returns the cleaned string.

        Args:
            input_string (str): Input string.
            specs (list): List of specialties.

        Returns:
            tuple: Tuple containing a list of specialties found and the cleaned string.
        """
        pattern = r'\((.*?)\)'
        results = re.findall(pattern, input_string)
        specs_in_word = []

        for result in results:
            cleaned_result = re.sub(r'[^a-zA-Zа-яА-ЯіІ ]', ' ', result)
            words = cleaned_result.split()
            founded_spec = []

            flag = True
            for word in words:
                for spec in specs:
                    if spec.lower().startswith(word):
                        founded_spec.append(spec)
                        break
                else:
                    flag = False
                    break
            if flag:
                specs_in_word.extend(founded_spec)
                input_string = input_string.replace(f'({result})', '')

        return specs_in_word, input_string.strip(', ').replace('  ', ' ')

    @staticmethod
    def _add_space_after_parenthesis(input_string):
        """
        Adds a space after closing parenthesis if not present.

        Args:
            input_string (str): Input string.

        Returns:
            str: String with spaces added after closing parenthesis.
        """
        new_string = ''
        for char in input_string:
            new_string += char
            if char == ')' and new_string[-2] != ' ':
                new_string += ' '
        return new_string

    @staticmethod
    def _get_groups(input_str: str):
        """
        Extracts groups from the input string.

        Args:
            input_str (str): Input string.

        Returns:
            list: List of groups.
        """
        if "лекція" in input_str.lower():
            return ["Лекція"]
        else:
            numbers = re.findall(r'\d+', input_str)
            return list(set(numbers))

    @staticmethod
    def _get_week_array(s):
        """
        Processes the input string to get an array of weeks.

        Args:
            s (str): Input string.

        Returns:
            list: List of weeks.
        """
        if '-' in s or ',' in s:
            result_array = []
            elements = s.split(',')

            for element in elements:
                if '-' in element:
                    start, end = map(int, element.split('-'))
                    result_array.extend(list(range(start, end + 1)))
                else:
                    result_array.append(int(element))
        else:
            result_array = [int(s)]

        return result_array

    @staticmethod
    def _proccess_time_range(input_str: str):
        """
        Processes the input string to get a dictionary of start, end, and full time.

        Args:
            input_str (str): Input time range string.

        Returns:
            dict: Dictionary containing start, end, and full time.
        """
        match = re.findall(r'\d+', input_str)

        start_time = f'{match[0]}:{match[1]}'
        end_time = f'{match[2]}:{match[3]}'

        time_dict = {
            'start': start_time,
            'end': end_time,
            'full': f'{start_time}-{end_time}'
        }

        return time_dict

    @staticmethod
    def _proccess_room(input_str: str):
        """
        Processes the room information.

        Args:
            input_str (str): Input room string.

        Returns:
            str: Processed room information.
        """
        if input_str.lower() == 'д':
            return 'Дистанційно'
        return input_str.capitalize()

    def _add_to_table(self, lesson_data, faculty, specs_of_faculty):
        """
        Adds data to the schedule table.

        Args:
            lesson_data (list): List containing day, time, subject, groups, weeks, and room information.
            faculty (str): Faculty name.
            specs_of_faculty (list): List of specialties.

        Returns:
            None
        """

        day, time, subject, groups, weeks, room = lesson_data

        # found the specialties to which the subject belongs
        # proccess name of the subject
        involved_specs, subject = self._get_specs_in_str(subject, specs_of_faculty)
        subject = self._add_space_after_parenthesis(subject)
        if not involved_specs:
            involved_specs = specs_of_faculty

        # proccess time
        time = self._proccess_time_range(time)
        # proccess room
        room = self._proccess_room(room)
        # proccess weeks
        weeks = self._get_week_array(weeks)
        # proccess group
        groups = self._get_groups(groups)

        for spec in involved_specs:
            if subject not in self._schedule_dict[faculty][spec]:
                self._schedule_dict[faculty][spec][subject] = {}

            for group in groups:
                self._schedule_dict[faculty][spec][subject][group] = {
                    "time": time,
                    "weeks": weeks,
                    "audience": room,
                    "day": day
                }

    def procces_excels(self, files: list):
        """
        Processes Excel files and populates the schedule dictionary.

        Args:
            files (list): List of Excel files to process.

        Returns:
            None
        """
        for file in files:
            try:
                df = pd.read_excel(file)
            except FileNotFoundError:
                print(f"Error: File {file} not found. Skipping...")
                continue
            faculty = None
            start_index = 0

            # get faculty name
            for index, row in df.iterrows():
                for column, value in row.items():
                    if not pd.isna(value):
                        value = str(value)
                        if value.startswith('Факультет'):
                            faculty = value
                            start_index = index
                            break
                else:
                    continue
                break

            self._schedule_dict[faculty] = {}

            # get specs names and study year
            specs = None
            study_year = None
            for index, row in df.iloc[start_index + 1:].iterrows():
                for column, value in row.items():
                    if not pd.isna(value):
                        value = str(value)
                        if value.startswith('Спеціальність'):
                            specs, study_year = self._extract_specialties_and_year(value)
                            start_index = index
                            break
                else:
                    continue
                break

            specs = [f'{spec}({study_year}р.н.)' for spec in specs]
            for spec in specs:
                self._schedule_dict[faculty][spec] = {}

            # skip everything to the beginning of the table
            for index, row in df.iloc[start_index + 1:].iterrows():
                for column, value in row.items():
                    if not pd.isna(value):
                        value = str(value)
                        if value.startswith('Понеділок'):
                            start_index = index
                            break
                else:
                    continue
                break

            day: str = 'Понеділок'
            time: str = '8.30-9.50'
            lesson_data = []
            EXPECTED_LESSON_DATA_LENGTH = 6
            for index, row in df.iloc[start_index:].iterrows():
                for column, value in row.items():
                    if not pd.isna(value):
                        # throw out all useless spaces
                        value = str(value).replace('\n', '').replace('  ', '')

                        # if the length is 0, then we check whether the day of the week is changed,
                        # if not, then we add the previous one
                        if len(lesson_data) == 0:
                            if value.replace('’', "`") in self.DAYS_NAME:
                                day = value
                            else:
                                lesson_data.append(day)

                        # if the length is 1, then we check whether the time has changed,
                        # if not, then we add the past
                        if len(lesson_data) == 1:
                            # print(value)
                            if self._is_valid_time_range(value):
                                time = value
                            else:
                                lesson_data.append(time)

                        # if the length is 2, then we expect the name of the lesson,
                        # the only thing that can come here is the name of the day of the week or the time
                        if len(lesson_data) == 2:
                            if value in self.DAYS_NAME:
                                lesson_data = []
                                day = value
                            elif self._is_valid_time_range(value):
                                lesson_data = [day]

                        # add the value to the lesson_data
                        lesson_data.append(value)

                        # add to the schedule the lesson_date
                        if len(lesson_data) == EXPECTED_LESSON_DATA_LENGTH:
                            self._add_to_table(lesson_data, faculty, specs)
                            lesson_data = []

    def save_to_json(self, output_file='schedule_output.json'):
        """
        Saves the schedule dictionary to a JSON file.

        Args:
            output_file (str): Output JSON file name.

        """
        try:
            with open(output_file, 'w', encoding='utf-16') as json_file:
                json.dump(self._schedule_dict, json_file, indent=4, ensure_ascii=False)
        except IOError:
            print(f"Error: Unable to write to file {output_file}.")

    def get_schedule_dict(self):
        """
        Returns the schedule dictionary.

        Returns:
            dict: Schedule dictionary.
        """
        return self._schedule_dict


def parse_arguments():
    """
    Function to parse command-line arguments for the script.

    Returns:
        argparse.Namespace: Parsed arguments.
    """
    parser = argparse.ArgumentParser(description='Process university schedules from Excel files.')

    parser.add_argument('files', metavar='file', type=str, nargs='+',
                        help='Excel file(s) to process')
    parser.add_argument('-o', '--output', metavar='output_file', type=str,
                        help='Output JSON file to save the processed schedule')

    return parser.parse_args()


if __name__ == "__main__":
    # Parse command-line arguments
    args = parse_arguments()

    # Use args.files to get the list of files
    schedule = Schedule(args.files)
    # Save to JSON file
    if args.output:
        schedule.save_to_json(args.output)
    else:
        schedule.save_to_json()

    print('Done')
