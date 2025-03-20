import camelot, fitz, ghostscript, matplotlib.pyplot as plt, os, pandas as pd, re

class _Service_Report:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def _open(self):
        self.doc = fitz.open(self.file_path)
        
    def _close(self):
        if self.doc is not None:
            self.doc.close()
            self.doc = None

    def _get_page(self, page_number: int) -> str:
        self._open()
        page = self.doc[page_number].get_text()
        self._close()
        return page

    def _extract_single_page_table(self, page_number: int, **kwargs) -> camelot.core.Table:
        tables = camelot.read_pdf(
            self.file_path,
            pages=str(page_number),
            **kwargs)

        if len(tables) > 0:
            return tables[0]
        else:
            raise ValueError(f"No table found in page n°{page_number}")
        
    def _convert_to_dataframe(self, table) -> pd.DataFrame:
        return table.df


    def plot_column_lines(self, table: camelot.core.Table, columns: list, title:str=None, kind='contour'):
        plt.figure()
        camelot.plot(table, kind=kind)
        for col in columns:
            plt.axvline(x=col, color='r', linestyle='--', label=f'Column {col}')
        plt.title(title)
        plt.xlabel('X-coordinate')
        plt.ylabel('Y-coordinate')
        plt.legend()
        plt.show()


    def _get_multiple_pages_table(self, starting_page_number: int, ending_page_number: int, **kwargs):
        tables = []
        
        for page_number in range(starting_page_number, ending_page_number + 1):
            try:
                object_table = self._extract_single_page_table(page_number, **kwargs)
                table = self._convert_to_dataframe(object_table)
                tables.append(table)
            except ValueError:
                continue
            
        if not tables:
            raise ValueError("No table found in document") 
        
        final_table = pd.concat(tables, ignore_index=True)
        return final_table
    
    def _clean_table(self, cleaning_pipeline) -> pd.DataFrame:
        pass
    
    def save_table_to_csv(self, table, name, folder_path):
        pass


