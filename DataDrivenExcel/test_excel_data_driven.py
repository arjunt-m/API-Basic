import pytest
import requests
from DataDrivenExcel.Library import Utils


class TestExcelDataDriven:

    base_url = "https://thetestingworldapi.com/api/studentsDetails"

    def test_all_rows(self):
        payloads = Utils.add_data_from_excel("Sheet1")
        for payload in payloads:
            print("Sending Payload:", payload)
            response = requests.post(self.base_url, json=payload)
            print("Status Code:", response.status_code)
            print("Response:", response.json())
            assert response.status_code == 201


    def test_specific_testcase(self):
        payloads = Utils.add_data_from_excel("Sheet1", testcase_id="TC02")
        for payload in payloads:
            print("Running TestCase TC02:", payload)
            response = requests.post(self.base_url, json=payload)
            print("Status Code:", response.status_code)
            print("Response:", response.json())
            assert response.status_code == 201


    def test_specific_row(self):
        payloads = Utils.add_data_from_excel("Sheet1", row_number=3)
        for payload in payloads:
            print("Running Row 3:", payload)
            response = requests.post(self.base_url, json=payload)
            print("Status Code:", response.status_code)
            print("Response:", response.json())
            assert response.status_code == 201