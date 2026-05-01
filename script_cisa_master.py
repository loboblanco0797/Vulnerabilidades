import pandas as pd
import requests
import logging
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# Configuración de logs profesionales
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class VulnerabilityDashboard:
    def __init__(self):
        self.url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
        self.df = None
        self.filename = f"Dashboard_CISA_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        self.font_header = Font(name="Calibri", size=18, bold=True, color="0d2a47")
        self.font_button_title = Font(name="Calibri", size=11, bold=True, color="0d2a47")
        self.border_box = Border(left=Side(style='medium'), right=Side(style='medium'), 
                                 top=Side(style='medium'), bottom=Side(style='medium'))

    def fetch_data(self):
        """Descarga e inteligencia de datos[cite: 1]."""
        try:
            logging.info("Sincronizando feed CISA KEV...")
            response = requests.get(self.url, timeout=15)
            response.raise_for_status()
            self.df = pd.DataFrame(response.json()['vulnerabilities'])
            
            # SOLUCIÓN AL ERROR: Convertir listas (como CWE) a texto plano[cite: 3]
            for col in self.df.columns:
                self.df[col] = self.df[col].apply(lambda x: ', '.join(map(str, x)) if isinstance(x, list) else x)
            
            self.df['dateAdded'] = pd.to_datetime(self.df['dateAdded']).dt.date
            logging.info(f"Datos cargados: {len(self.df)} registros.")
            return True
        except Exception as e:
            logging.error(f"Falla en datos: {e}")
            return False

    def setup_dashboard_sheet(self, wb):
        """Diseña la interfaz interactiva."""
        ws = wb.active
        ws.title = "Dashboard"
        ws.sheet_view.showGridLines = False 
        
        ws.merge_cells("C2:M3")
        title_cell = ws["C2"]
        title_cell.value = "SISTEMA AUTOMATIZADO DE INTELIGENCIA DE AMENAZAS"
        title_cell.font = self.font_header
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        buttons = [
            ("MICROSOFT", "Análisis de fallos en Windows/Azure.", "Vulnerabilidades_Microsoft"),
            ("CISCO", "Amenazas en Networking y Routers.", "Vulnerabilidades_Cisco"),
            ("APPLE", "Riesgos en macOS y iOS.", "Vulnerabilidades_Apple"),
            ("CLOUD", "Vulnerabilidades en AWS/GCP/Azure.", "Vulnerabilidades_Cloud"),
            ("OPEN SOURCE", "Librerías críticas (Log4j, Apache).", "Vulnerabilidades_OSS"),
            ("OTROS", "Otros fabricantes (Adobe, Fortinet).", "Vulnerabilidades_Otros")
        ]
        
        start_row, start_col = 5, 3
        for index, (title, desc, sheet_target) in enumerate(buttons):
            r, c = start_row + (index // 3) * 6, start_col + (index % 3) * 4
            main_cell = ws.cell(row=r, column=c, value=f"{title}\n{desc}")
            main_cell.font = self.font_button_title
            main_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            main_cell.border = self.border_box
            main_cell.hyperlink = f"#{sheet_target}!A1"
            ws.merge_cells(f"{get_column_letter(c)}{r}:{get_column_letter(c+2)}{r+3}")
        
        for col_num in range(3, 14):
            ws.column_dimensions[get_column_letter(col_num)].width = 20

    def add_data_sheets(self, wb):
        """Crea hojas categorizadas con formato de tabla[cite: 1, 3]."""
        categories = [
            ("Vulnerabilidades_Microsoft", self.df['vendorProject'].str.contains('Microsoft', case=False)),
            ("Vulnerabilidades_Cisco", self.df['vendorProject'].str.contains('Cisco', case=False)),
            ("Vulnerabilidades_Apple", self.df['vendorProject'].str.contains('Apple', case=False)),
            ("Vulnerabilidades_Cloud", self.df['vendorProject'].str.contains('AWS|Cloud|Google|Azure', case=False)),
            ("Vulnerabilidades_OSS", self.df['shortDescription'].str.contains('Open Source|Librer|OpenSSL', case=False)),
            ("Vulnerabilidades_Otros", ~self.df['vendorProject'].str.contains('Microsoft|Cisco|Apple|AWS|Cloud', case=False))
        ]
        
        for sheet_name, mask in categories:
            df_filtered = self.df[mask]
            if not df_filtered.empty:
                ws = wb.create_sheet(title=sheet_name)
                ws.append(list(df_filtered.columns))
                for row in df_filtered.values:
                    ws.append(list(row))
                
                full_range = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                # Limpiar nombre de tabla para Excel[cite: 3]
                clean_name = sheet_name.replace("_", "")
                tab = Table(displayName=f"Tabla{clean_name}", ref=full_range)
                tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                ws.add_table(tab)
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 30

    def create_dashboard(self):
        """Genera y abre el Dashboard[cite: 3]."""
        wb = Workbook()
        self.setup_dashboard_sheet(wb)
        self.add_data_sheets(wb)
        try:
            wb.save(self.filename)
            logging.info("Dashboard interactivo generado exitosamente.")
            os.startfile(self.filename)
        except Exception as e:
            logging.error(f"Error al guardar: {e}")

if __name__ == "__main__":
    app = VulnerabilityDashboard()
    if app.fetch_data():
        app.create_dashboard()