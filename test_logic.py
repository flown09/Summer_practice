import unittest
import pandas as pd
from logic import read_data, apply_conditions

class TestFileComparisonLogic(unittest.TestCase):

    def setUp(self):
        self.df1 = pd.DataFrame({
            'Имя': ['Аня', 'Борис', 'Виктор'],
            'ИНН': ['111', '222', '333']
        })
        self.df2 = pd.DataFrame({
            'Имя': ['Аня', 'Глеб'],
            'ИНН': ['111', '444']
        })

    def test_apply_conditions_match(self):
        conditions = [('Имя', 'Совпадают')]
        result = apply_conditions(self.df1, self.df2, conditions)
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]['Имя'], 'Аня')

    def test_apply_conditions_not_match(self):
        conditions = [('ИНН', 'Не совпадают')]
        result = apply_conditions(self.df1, self.df2, conditions)
        self.assertEqual(len(result), 2)
        self.assertNotIn('111', result['ИНН'].values)

    def test_invalid_field(self):
        conditions = [('Возраст', 'Совпадают')]
        with self.assertRaises(ValueError):
            apply_conditions(self.df1, self.df2, conditions)

    def test_invalid_condition_type(self):
        conditions = [('Имя', 'Содержит')]
        with self.assertRaises(ValueError):
            apply_conditions(self.df1, self.df2, conditions)

if __name__ == '__main__':
    unittest.main()
