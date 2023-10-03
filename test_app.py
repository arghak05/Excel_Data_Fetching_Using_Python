import unittest
from refactored_app import *

class TestYourFunctions(unittest.TestCase):

    def setUp(self):
        self.wb = openpyxl.load_workbook('Excel_Data_Reader/spain_data.xlsx')
        self.sheet = self.wb['Hoja2']

    def test_find_start_and_end_rows(self):
        start_row, end_row = find_start_and_end_rows(self.sheet)
        self.assertEqual(start_row, 13)
        self.assertEqual(end_row, 15)

    def test_identify_weekday_and_weekend_columns(self):
        weekday_col_start, weekend_col_start = identify_weekday_and_weekend_columns(self.sheet)
        self.assertEqual(weekday_col_start, 4)
        self.assertEqual(weekend_col_start, 54)

    def test_extract_station_names(self):
        mock_data = ['TOTAL RADIO DÍA DE AYER', 'RADIO OM DÍA DE AYER']
        start_row, end_row = 13, 15
        weekday_col_start = 4
        station_names = extract_station_names(self.sheet, start_row, end_row, weekday_col_start)
        self.assertEqual(station_names, mock_data)

    def test_extract_time_stamps(self):
        mock_data = ['06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00']
        start_row = 13
        weekday_col_start = 4
        time_stamps = extract_time_stamps(self.sheet, start_row, weekday_col_start)
        self.assertEqual(time_stamps[0:4], mock_data)

    def test_process_data(self):
        mock_data = {'Tv_Program_Channel': 'TOTAL RADIO DÍA DE AYER', 'Time': {'6.00': 3960.0, '7.00': 9718.0, '8.00': 13806.0, '9.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0}, 'Flag': 0}
        start_row, end_row = 13, 15
        weekday_col_start = 4
        weekend_col_start = 54
        station_name = ['TOTAL RADIO DÍA DE AYER', 'RADIO OM DÍA DE AYER']
        time_stamp = ['06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00', '08.00 a 08.30', '08.30 a 09.00', '09.00 a 09.30', '09.30 a 10.00', '10.00 a 10.30', '10.30 a 11.00', '11.00 a 11.30', '11.30 a 12.00', '12.00 a 12.30', '12.30 a 13.00']
        nested_aud = process_data(self.sheet, start_row, end_row, weekday_col_start, weekend_col_start, station_name, time_stamp)
        self.assertEqual(nested_aud[0], mock_data)

    def test_generate_sql_insert_statements(self):
        nested_aud = [{'Tv_Program_Channel': 'TOTAL RADIO DÍA DE AYER', 'Time': {'6.00': 3960.0, '7.00': 9718.0, '8.00': 13806.0, '9.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0}, 'Flag': 0}]
        day_mapping = {0: "Monday", 1: "Tuesday", 2: "Wednesday", 3: "Thursday", 4: "Friday", 5: "Saturday", 6: "Sunday"}
        mock_data = ["INSERT INTO radio_audience(station_name, timestamp, audience, flag, day_of_week) VALUES (TOTAL RADIO DÍA DE AYER,{'6.00': 3960.0, '7.00': 9718.0, '8.00': 13806.0, '9.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0},0,Monday);"]
        sql_statements = generate_sql_insert_statements(nested_aud, day_mapping)
        self.assertEqual(sql_statements[0],mock_data[0])

if __name__ == '__main__':
    unittest.main()
