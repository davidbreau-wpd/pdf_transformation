import camelot, fitz, ghostscript, matplotlib.pyplot as plt, os, pandas as pd, re
from .service_report import _Service_Report

class Vestas_Report(_Service_Report):
    def __init__(self, file_path):
        super().__init__(file_path)
        
        self.columns = [65, 330, 350]
        self.camelot_params = {
            'flavor': 'stream',
            'edge_tol': 500,
            'row_tol': 10,
            'columns': [",".join(str(col) for col in self.columns)]
        } 
    
    def find_starting_page(self) -> int:
        self._open()
        found_page = None
        for page_num in range(len(self.doc)):
            page_text = self.doc[page_num].get_text()
            if "Service Inspection Form" in page_text:
                found_page = page_num
                break
        self._close()
        
        if found_page is not None:
            return found_page
        raise ValueError("No 'Service Inspection Form' page found in document")
        
    def get_header_informations(self):
        first_page = super()._get_page(0)
        
        turbine_number = re.search(r'Turbine No\./Id:\s*(\d+)', first_page).group(1) if re.search(r'Turbine No\./Id:\s*(\d+)', first_page) else None
        service_order = re.search(r'Service Order:\s*(\d+)', first_page).group(1) if re.search(r'Service Order:\s*(\d+)', first_page) else None
        pad_no = (match.group(1).strip() if (match := re.search(r'PAD No\.\s*([^\n]+)', first_page)) else None)
        turbine_type = re.search(r'Turbine Type:\s*([\w\d]+)', first_page).group(1) if re.search(r'Turbine Type:\s*([\w\d]+)', first_page) else None
        start_date = re.search(r'Start Date:\s*([\d\.]+)', first_page).group(1) if re.search(r'Start Date:\s*([\d\.]+)', first_page) else None
        end_date = (match.group(1) if (match := re.search(r'End Date:\s*([\d\.]+)', first_page)) else None)
        date_and_time_of_receipt = (match.group(1).strip() if (match := re.search(r'Date & Time of Receipt\s*([\d\.\s:]+)', first_page)) else None)
        reason_for_call_out = (match.group(1) if (match := re.search(r'Reason for Call Out:\s*([^\n]+)', first_page)) else None)
                
        customer_address = ([line.strip() for line in match.group(1).split('\n') if line.strip()] 
                       if (match := re.search(r"Customer's Address:\s*(.*?)Site's Address:", first_page, re.DOTALL)) 
                       else None)
        
        header_informations = {
            'turbine_number': turbine_number,
            'service_order': service_order,
            'pad_no': pad_no,
            'turbine_type': turbine_type,
            'start_date': start_date,
            'customer_address': customer_address,
            'date_and_time_of_receipt': date_and_time_of_receipt,
            'reason_for_call_out': reason_for_call_out
        }
        
        return header_informations
    
    def check_column_lines(self, page_number: int, columns=None):
        if columns is None:
            columns = self.columns
        table = super()._extract_single_page_table(
            page_number=page_number,
            **self.camelot_params
        )
        super().plot_column_lines(table,columns)
        
    def _extract_full_table(self, columns=None) -> pd.DataFrame:
        starting_page = self.find_starting_page() + 1
        self._open()
        ending_page = len(self.doc)
        self._close()
        
        params = self.camelot_params.copy()
        if columns is not None:
            params['columns'] = [",".join(str(col) for col in columns)]
            
        return self._get_multiple_pages_table(
            starting_page_number=starting_page,
            ending_page_number=ending_page,
            **params
        )
        