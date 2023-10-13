import argparse
import json
import re
import pandas as pd


class Schedule:
    DAYS_NAME = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', 'П`ятниця', 'Субота', 'Неділя']
    _schedule_dict = {}

    def __init__(self, files: list | str):
        if isinstance(files, str):
            files: list = [files]
        self.procces_excels(files)

    @staticmethod
    def _extract_specialties_and_year(text):
        """
        Extracts specialties and study year from the given text.

        Args:
        - text (str): Input text containing specialties and study year information.

        Returns:
        - tuple: A tuple containing a list of specialties and the study year.
        """
        specialty_matches = re.findall(r'[«"](.*?)[»"]', text)
        year_matches = re.findall(r'(\d+) р\.н\.', text)

        specialties = [spec.strip() for spec in specialty_matches]
        years = int(year_matches[0])

        return specialties, years

    @staticmethod
    def _is_valid_time_range(input_str: str):
        """
        Checks if the input string represents a valid time range.

        Args:
        - input_str (str): Input string to be checked.

        Returns:
        - bool: True if the input string is a valid time range, False otherwise.
        """
        pattern = re.compile(r'^\d{1,2}[.:]\d{2}-\d{1,2}[.:]\d{2}$')
        return bool(pattern.match(input_str))

    @staticmethod
    def _get_specs_in_str(input_string, specs):
        """
        Extracts specialties mentioned in a string based on a given list of specialties.

        Args:
        - input_string (str): Input string containing information about specialties.
        - specs (list): List of available specialties.

        Returns:
        - tuple: A tuple containing a list of extracted specialties and the modified input string.
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
        Adds a space after each closing parenthesis in the input string.

        Args:
        - input_string (str): Input string.

        Returns:
        - str: Modified input string with spaces added after closing parentheses.
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
        - input_str (str): Input string containing information about groups.

        Returns:
        - list: List of extracted group numbers.
        """
        if "лекція" in input_str.lower():
            return ["Лекція"]
        else:
            numbers = re.findall(r'\d+', input_str)
            return list(set(numbers))

    @staticmethod
    def _get_week_array(s):
        """
        Parses a string representing weeks and returns an array of week numbers.

        Args:
        - s (str): Input string representing weeks.

        Returns:
        - list: List of week numbers.
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
        Processes a string representing a time range and returns a dictionary with start, end, and full time.

        Args:
        - input_str (str): Input string representing a time range.

        Returns:
        - dict: Dictionary containing start time, end time, and full time.
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
        Processes a string representing a room and returns the processed room information.

        Args:
        - input_str (str): Input string representing a room.

        Returns:
        - str: Processed room information.
        """
        if input_str.lower() == 'д':
            return 'Дистанційно'
        return input_str.capitalize()

    def _add_to_table(self, lesson_data, faculty, specs_of_faculty):
        """
        Adds a lesson to the schedule table for a specific faculty and specialties.

        Args:
        - lesson_data (list): List containing information about the lesson (day, time, subject, groups, weeks, room).
        - faculty (str): Name of the faculty for which the lesson is being added.
        - specs_of_faculty (list): List of specialties associated with the faculty.
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
        Processes Excel files to populate the _schedule_dict attribute with schedule data.

        Args:
        - files (list): List of Excel file names.
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
        Saves the processed schedule data to a JSON file.

        Args:
        - output_file (str): Name of the output JSON file.
        """
        try:
            with open(output_file, 'w', encoding='utf-16') as json_file:
                json.dump(self._schedule_dict, json_file, indent=4, ensure_ascii=False)
        except IOError:
            print(f"Error: Unable to write to file {output_file}.")

    def get_schedule_dict(self):
        return self._schedule_dict


def parse_arguments():
    """
    Parses command-line arguments.

    Returns:
    - argparse.Namespace: An object containing parsed arguments.
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
