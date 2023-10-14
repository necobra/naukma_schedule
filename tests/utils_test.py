# test_functions.py
import unittest
from schedule.utils import *


class TestYourModuleFunctions(unittest.TestCase):

    def test_extract_specialties_and_year(self):
        test_cases = [
            {
                'input': 'Спеціальність «Економіка», «Фінанси, банківська справа та страхування», «Маркетинг», '
                         '«Менеджмент» 3 р.н.',
                'expected_output': (
                    ['Економіка', 'Фінанси, банківська справа та страхування', 'Маркетинг', 'Менеджмент'], 3)
            },
            {
                'input': 'Спеціальність "Інженерія програмного забезпечення ", 3 р.н.',
                'expected_output': (['Інженерія програмного забезпечення'], 3)
            },
        ]

        for case in test_cases:
            specialties, year = extract_specialties_and_year(case['input'])
            self.assertEqual(specialties, case['expected_output'][0])
            self.assertEqual(year, case['expected_output'][1])

    def test_proccess_time_range(self):
        test_cases = [
            {'input': '08:30-10:00',
             'expected_output': {'start': '08:30', 'end': '10:00', 'full': '08:30-10:00'}},
            {'input': '08.30-10.00',
             'expected_output': {'start': '08:30', 'end': '10:00', 'full': '08:30-10:00'}},
            {'input': '08:30-10.00',
             'expected_output': {'start': '08:30', 'end': '10:00', 'full': '08:30-10:00'}},
            {'input': '8:30-10.00',
             'expected_output': {'start': '08:30', 'end': '10:00', 'full': '08:30-10:00'}},
        ]

        for case in test_cases:
            result = proccess_time_range(case['input'])
            self.assertEqual(result, case['expected_output'])

    def test_proccess_room(self):
        test_cases = [
            {'input': '6-301', 'expected_output': '6-301'},
            {'input': 'д', 'expected_output': 'Дистанційно'},
            {'input': 'Д', 'expected_output': 'Дистанційно'},
            {'input': 'дистанційно', 'expected_output': 'Дистанційно'},
        ]

        for case in test_cases:
            result = proccess_room(case['input'])
            self.assertEqual(result, case['expected_output'])

    def test_get_week_array(self):
        test_cases = [
            {'input': '1-5,7,10', 'expected_output': [1, 2, 3, 4, 5, 7, 10]},
            {'input': '15', 'expected_output': [15]},
            {'input': '2,3,7,10,11,12', 'expected_output': [2, 3, 7, 10, 11, 12]},
            {'input': '1-4,7-10', 'expected_output': [1, 2, 3, 4, 7, 8, 9, 10]},
        ]

        for case in test_cases:
            result = get_week_array(case['input'])
            self.assertEqual(result, case['expected_output'])

    def test_get_specs_in_str(self):
        test_cases = [
            {
                'input': 'Навчально-науковий семінар з економіки (екон.) проф. Бураковський І.В.',
                'specs': ['Економіка', 'Фінанси, банківська справа та страхування', 'Маркетинг', 'Менеджмент'],
                'expected_output': (['Економіка'], 'Навчально-науковий семінар з економіки проф. Бураковський І.В.')
            },
            {
                'input': 'Гроші та кредит (фін.+мар.) ',
                'specs': ['Економіка', 'Фінанси, банківська справа та страхування', 'Маркетинг', 'Менеджмент'],
                'expected_output': (['Фінанси, банківська справа та страхування', 'Маркетинг'], 'Гроші та кредит'),
            },
            {
                'input': 'Економіко-математичне моделювання-ІІ (Економетрика) (фін.) ст. викл. Дадашова П.А.',
                'specs': ['Економіка', 'Фінанси, банківська справа та страхування', 'Маркетинг', 'Менеджмент'],
                'expected_output': (['Фінанси, банківська справа та страхування'],
                                    'Економіко-математичне моделювання-ІІ (Економетрика) ст. викл. Дадашова П.А.'),
            }
        ]

        for case in test_cases:
            result, cleaned_string = get_specs_in_str(case['input'], case['specs'])
            self.assertEqual(result, case['expected_output'][0])
            self.assertEqual(cleaned_string, case['expected_output'][1])

    def test_get_groups(self):
        test_cases = [
            {'input': '5ф', 'expected_output': [5]},
            {'input': 'лекція', 'expected_output': ['Лекція']},
            {'input': '1 2 3 4', 'expected_output': [1, 2, 3, 4]},
            {'input': 'Лекція 1п', 'expected_output': ['Лекція']},
            {'input': '2 ф+мар', 'expected_output': [2]},
        ]

        for case in test_cases:
            result = get_groups(case['input'])
            self.assertEqual(result, case['expected_output'])


if __name__ == '__main__':
    unittest.main()
