"""
This module provides functions for extracting and processing information related to schedule
"""
import re
from typing import List, Tuple, Union


def extract_specialties_and_year(text: str) -> Tuple[List[str], int]:
    """
    Extracts specialties and study year from a text.

    Args:
        text (str): Input text.

    Returns:
        Tuple: List of specialties and the study year.
    """
    specialty_matches = re.findall(r'[«"](.*?)[»"]', text)
    year_matches = re.findall(r'(\d+) р\.н\.', text)

    specialties = [spec.strip() for spec in specialty_matches]
    years = int(year_matches[0]) if year_matches else 0

    return specialties, years


def is_valid_time_range(input_str: str) -> bool:
    """
    Checks if the input string is a valid time range.

    Args:
        input_str (str): Input time range string.

    Returns:
        bool: True if the input is a valid time range, False otherwise.
    """
    pattern = re.compile(r'^\d{1,2}[.:]\d{2}-\d{1,2}[.:]\d{2}$')
    return bool(pattern.match(input_str))


def get_specs_in_str(input_string: str, specs: List[str]) -> Tuple[List[str], str]:
    """
    Extracts specialties mentioned in a string and returns the cleaned string.

    Args:
        input_string (str): Input string.
        specs (List): List of specialties.

    Returns:
        Tuple: List of specialties found and the cleaned string.
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


def add_space_after_parenthesis(input_string: str) -> str:
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


def get_groups(input_str: str) -> List[Union[str, int]]:
    """
    Extracts groups from the input string.

    Args:
        input_str (str): Input string.

    Returns:
        List: List of groups.
    """
    if "лекція" in input_str.lower():
        return ["Лекція"]
    else:
        numbers = re.findall(r'\d+', input_str)
        return list(set(int(num) for num in numbers))


def get_week_array(s: str) -> List[int]:
    """
    Processes the input string to get an array of weeks.

    Args:
        s (str): Input string.

    Returns:
        List: List of weeks.
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


def proccess_time_range(input_str: str) -> dict:
    """
    Processes the input string to get a dictionary of start, end, and full time.

    Args:
        input_str (str): Input time range string.

    Returns:
        dict: Dictionary containing start, end, and full time.
    """
    match = re.findall(r'\d+', input_str)

    if len(match[0]) == 1:
        match[0] = f'0{match[0]}'

    start_time = f'{match[0]}:{match[1]}'
    end_time = f'{match[2]}:{match[3]}'

    time_dict = {
        'start': start_time,
        'end': end_time,
        'full': f'{start_time}-{end_time}'
    }

    return time_dict


def proccess_room(input_str: str) -> str:
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
