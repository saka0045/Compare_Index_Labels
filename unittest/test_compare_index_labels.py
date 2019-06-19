from unittest import TestCase
from compare_index_labels import *
from ErrorLib import *
import os

current_directory = os.getcwd()

class Test_parse_index_sheet(TestCase):
    def setUp(self):
        self.index_sheet_path = current_directory + "/Test_File_A.csv"
        self.index_sheet_path_error = current_directory + "/Test_File_B.csv"
        self.expected_index_information = {'1': 'AAAAAAAAAA', '2': 'BBBBBBBBBB', '3': 'CCCCCCCCCC',
                                           '4': 'DDDDDDDDDD', '5': 'EEEEEEEEEE'}

    def test_parse_index_sheet_happy(self):
        index_information = parse_index_sheet(self.index_sheet_path)
        self.assertEqual(self.expected_index_information, index_information, msg="index information")

    def test_parse_index_sheet_error(self):
        with self.assertRaises(DuplicateIndexError):
            parse_index_sheet(self.index_sheet_path_error)
