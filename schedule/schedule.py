"""
Module for processing university schedules from Excel files.

Classes:
    Schedule: Class for processing and storing university schedules.
"""
import json
import pandas as pd
from .utils import *


class Schedule:
    """
    Class for processing and storing university schedules.

    Attributes:
        DAYS_NAME (list): List of day names.
        _schedule_dict (dict): Dictionary to store the processed schedule.

    Methods:
        __init__: Initializes the Schedule class.
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
        involved_specs, subject = get_specs_in_str(subject, specs_of_faculty)
        subject = add_space_after_parenthesis(subject)
        if not involved_specs:
            involved_specs = specs_of_faculty

        # proccess time
        time = proccess_time_range(time)
        # proccess room
        room = proccess_room(room)
        # proccess weeks
        weeks = get_week_array(weeks)
        # proccess group
        groups = get_groups(groups)

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
                            specs, study_year = extract_specialties_and_year(value)
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
                            if is_valid_time_range(value):
                                time = value
                            else:
                                lesson_data.append(time)

                        # if the length is 2, then we expect the name of the lesson,
                        # the only thing that can come here is the name of the day of the week or the time
                        if len(lesson_data) == 2:
                            if value in self.DAYS_NAME:
                                lesson_data = []
                                day = value
                            elif is_valid_time_range(value):
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
                print(f'The data is saved to file {output_file}.')
        except IOError:
            print(f"Error: Unable to write to file {output_file}.")

    def get_schedule_dict(self):
        """
        Returns the schedule dictionary.

        Returns:
            dict: Schedule dictionary.
        """
        return self._schedule_dict
