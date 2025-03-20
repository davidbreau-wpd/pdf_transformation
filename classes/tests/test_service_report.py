import pytest
from unittest.mock import MagicMock
import pandas as pd
from service_report import _Service_Report

class TestServiceReport:
    def test_open_and_close_pdf_document(self, mocker):
        mock_open = mocker.patch('fitz.open', return_value=MagicMock())
        report = _Service_Report("dummy_path.pdf")
        report._open()
        report._close()
        mock_open.assert_called_once_with("dummy_path.pdf")
        assert report.doc is None

    def test_extract_single_page_table(self, mocker):
        mock_read_pdf = mocker.patch('camelot.read_pdf', return_value=[MagicMock()])
        report = _Service_Report("dummy_path.pdf")
        table = report._extract_single_page_table(1)
        mock_read_pdf.assert_called_once_with("dummy_path.pdf", pages='1')
        assert table is not None

    def test_convert_table_to_dataframe(self):
        mock_table = MagicMock()
        mock_table.df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        report = _Service_Report("dummy_path.pdf")
        df = report._convert_to_dataframe(mock_table)
        pd.testing.assert_frame_equal(df, mock_table.df)

    def test_handle_non_existent_page_number(self, mocker):
        mock_open = mocker.patch('fitz.open', return_value=MagicMock())
        mock_open.return_value.__getitem__.side_effect = IndexError
        report = _Service_Report("dummy_path.pdf")
        with pytest.raises(IndexError):
            report._get_page(100)

    def test_raise_error_no_table_found(self, mocker):
        mock_read_pdf = mocker.patch('camelot.read_pdf', return_value=[])
        report = _Service_Report("dummy_path.pdf")
        with pytest.raises(ValueError, match="No table found in page n°1"):
            report._extract_single_page_table(1)

    def test_concatenate_tables_multiple_pages(self, mocker):
        mock_read_pdf = mocker.patch('camelot.read_pdf', side_effect=[
            [MagicMock(df=pd.DataFrame({'A': [1], 'B': [2]}))],
            [],
            [MagicMock(df=pd.DataFrame({'A': [3], 'B': [4]}))]
        ])
        report = _Service_Report("dummy_path.pdf")
        df = report._extract_multiple_pages_table(1, 3)
        expected_df = pd.DataFrame({'A': [1, 3], 'B': [2, 4]})
        pd.testing.assert_frame_equal(df, expected_df)