import camelot, fitz, ghostscript, matplotlib.pyplot as plt, os, pandas as pd, re

class PDF:
    def __init__(self, name, path=None):
        self.path = path
        self.name = name
        # Ouvrir le document (variable locale)
        file_path = os.path.join(self.path, self.name) if self.path else self.name
        opened_doc = fitz.open(file_path)
        self.page_numbers = len(opened_doc)
        
        # Obtenir toutes les informations nécessaires
        self.header_informations = self._get_header_informations(opened_doc)
        self.sif_page = self._get_sif_page(opened_doc)
        
        # Fermer le document
        opened_doc.close()
    
    def _get_header_informations(self, opened_doc):
        text = opened_doc[0].get_text()
       
        turbine_number = re.search(r'Turbine No\./Id:\s*(\d+)', text).group(1) if re.search(r'Turbine No\./Id:\s*(\d+)', text) else None
        service_order = re.search(r'Service Order:\s*(\d+)', text).group(1) if re.search(r'Service Order:\s*(\d+)', text) else None
        pad_no = (match.group(1).strip() if (match := re.search(r'PAD No\.\s*([^\n]+)', text)) else None)
        turbine_type = re.search(r'Turbine Type:\s*([\w\d]+)', text).group(1) if re.search(r'Turbine Type:\s*([\w\d]+)', text) else None
        start_date = re.search(r'Start Date:\s*([\d\.]+)', text).group(1) if re.search(r'Start Date:\s*([\d\.]+)', text) else None
        end_date = (match.group(1) if (match := re.search(r'End Date:\s*([\d\.]+)', text)) else None)
        date_and_time_of_receipt = (match.group(1).strip() if (match := re.search(r'Date & Time of Receipt\s*([\d\.\s:]+)', text)) else None)
        reason_for_call_out = (match.group(1) if (match := re.search(r'Reason for Call Out:\s*([^\n]+)', text)) else None)
                
        customer_address = ([line.strip() for line in match.group(1).split('\n') if line.strip()] 
                       if (match := re.search(r"Customer's Address:\s*(.*?)Site's Address:", text, re.DOTALL)) 
                       else None)

        return {
            'turbine_number': turbine_number,
            'service_order': service_order,
            'pad_no': pad_no,
            'turbine_type': turbine_type,
            'start_date': start_date,
            'customer_address': customer_address,
            'date_and_time_of_receipt': date_and_time_of_receipt,
            'reason_for_call_out': reason_for_call_out
        }
    
    def _get_sif_page(self, opened_doc):
        """Trouve la page contenant 'Service Inspection Form'"""
        for page_num in range(self.page_numbers):
            if "Service Inspection Form" in opened_doc[page_num].get_text():
                return page_num
        raise ValueError("'Service Inspection Form' non trouvé dans le document")
    
    def get_page_table(self, page_number):
        tables = camelot.read_pdf(
            self.path,
            pages=str(page_number),
            flavor='stream',
            edge_tol=500,
            row_tol=10,
            columns=['65,330,350'] 
        )
        
        if len(tables) > 0:
            return tables[0].df
        else:
            raise ValueError(f"Aucune table trouvée à la page {page_number}")
    
    def get_full_table(self):
        # Liste pour stocker tous les DataFrames
        all_tables = []
        
        # Extraire les tables de chaque page, de sif_page jusqu'à la fin
        for page_num in range(self.sif_page, self.page_numbers):
            try:
                # +1 car Camelot commence à 1
                table = self.get_page_table(page_num + 1)
                all_tables.append(table)
            except ValueError:
                continue
        
        if not all_tables:
            raise ValueError("Aucune table trouvée dans le document")
        
        final_table = pd.concat(all_tables, ignore_index=True)
        return final_table
    
    def save_csv(self, path, name):
        # Créer le dossier s'il n'existe pas
        if not os.path.exists(path):
            os.makedirs(path)
        self.get_full_table().to_csv(f"{path}/{name}.csv", index=False)
