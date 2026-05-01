import pandas as pd
import requests
import logging
import os
from datetime import datetime
from openpyxl import Workbook
# Se añade PatternFill para el manejo de colores de fondo
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill 
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# Configuración de logs profesionales
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class VulnerabilityDashboard:
    def __init__(self):
        self.url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
        self.df = None
        self.filename = f"Dashboard_SOC_CISA_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        # Estilos base
        self.font_header = Font(name="Segoe UI", size=20, bold=True, color="FFFFFF")
        self.font_button_title = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

    def fetch_data(self):
        """Descarga e inteligencia de datos."""
        try:
            logging.info("Sincronizando feed CISA KEV...")
            response = requests.get(self.url, timeout=15)
            response.raise_for_status()
            self.df = pd.DataFrame(response.json()['vulnerabilities'])
            
            # Sanitización: Convertir listas a texto plano para compatibilidad con Excel
            for col in self.df.columns:
                self.df[col] = self.df[col].apply(lambda x: ', '.join(map(str, x)) if isinstance(x, list) else x)
            
            self.df['dateAdded'] = pd.to_datetime(self.df['dateAdded']).dt.date
            logging.info(f"Datos cargados: {len(self.df)} registros.")
            return True
        except Exception as e:
            logging.error(f"Falla en obtención de datos: {e}")
            return False

    def setup_dashboard_sheet(self, wb):
        """Diseña la interfaz interactiva con estética SOC (Dark Mode)."""
        ws = wb.active
        ws.title = "Dashboard"
        ws.sheet_view.showGridLines = False 
        
        # Paleta de colores SOC
        bg_dark = "0B192C"      # Fondo oscuro
        btn_blue = "1E3A8A"     # Azul corporativo
        text_white = "FFFFFF"   

        # 1. Aplicar fondo oscuro a la interfaz
        for r in range(1, 40):
            for c in range(1, 20):
                ws.cell(row=r, column=c).fill = PatternFill(start_color=bg_dark, end_color=bg_dark, fill_type="solid")

        # 2. Título Principal
        ws.merge_cells("C2:M3")
        title_cell = ws["C2"]
        title_cell.value = "CENTRO DE INTELIGENCIA DE AMENAZAS - CISA KEV"
        title_cell.font = self.font_header
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 3. Configuración de Cards (Botones)
        buttons = [
            ("MICROSOFT", "Sistemas Windows & Server", "Vulnerabilidades_Microsoft", "4F81BD"),
            ("CISCO", "Infraestructura de Red", "Vulnerabilidades_Cisco", "3A86FF"),
            ("APPLE", "Ecosistema iOS & macOS", "Vulnerabilidades_Apple", "8E9AAF"),
            ("CLOUD", "Entornos AWS/GCP/Azure", "Vulnerabilidades_Cloud", "00B4D8"),
            ("OPEN SOURCE", "Librerías y Core OSS", "Vulnerabilidades_OSS", "FB8500"),
            ("OTROS", "Fabricantes Terceros", "Vulnerabilidades_Otros", "8338EC")
        ]
        
        start_row, start_col = 6, 3
        
        for index, (title, desc, sheet_target, accent_color) in enumerate(buttons):
            r = start_row + (index // 3) * 7
            c = start_col + (index % 3) * 4
            
            # Asignar valor y estilo a la celda principal antes de fusionar
            main_cell = ws.cell(row=r, column=c)
            main_cell.value = f"{title}\n\n{desc}"
            main_cell.font = self.font_button_title
            main_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            main_cell.hyperlink = f"#{sheet_target}!A1"
            
            # Definir borde de la Card con el color de acento
            card_border = Border(left=Side(style='thick', color=accent_color), 
                                 right=Side(style='thick', color=accent_color), 
                                 top=Side(style='thick', color=accent_color), 
                                 bottom=Side(style='thick', color=accent_color))
            
            # Aplicar relleno y bordes a las celdas del rango de la Card
            for row_btn in range(r, r + 5):
                for col_btn in range(c, c + 3):
                    cell = ws.cell(row=row_btn, column=col_btn)
                    cell.fill = PatternFill(start_color=btn_blue, end_color=btn_blue, fill_type="solid")
                    cell.border = card_border
            
            # Fusionar celdas de la Card
            range_str = f"{get_column_letter(c)}{r}:{get_column_letter(c+2)}{r+4}"
            ws.merge_cells(range_str)
        
        # Ajuste de ancho de columnas
        for col_num in range(3, 14):
            ws.column_dimensions[get_column_letter(col_num)].width = 22

    def add_data_sheets(self, wb):
        """Crea hojas de datos categorizadas con formato de tabla profesional."""
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
                # Agregar encabezados
                ws.append(list(df_filtered.columns))
                # Agregar registros
                for row in df_filtered.values:
                    ws.append(list(row))
                
                # Convertir a Tabla oficial de Excel
                full_range = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                clean_name = sheet_name.replace("_", "")
                tab = Table(displayName=f"Tabla{clean_name}", ref=full_range)
                tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                ws.add_table(tab)
                
                # Auto-ajuste de columnas
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 35

    def create_dashboard(self):
        """Orquesta la creación y apertura del Dashboard."""
        wb = Workbook()
        self.setup_dashboard_sheet(wb)
        self.add_data_sheets(wb)
        try:
            wb.save(self.filename)
            logging.info("¡Dashboard SOC Premium generado exitosamente!")
            os.startfile(self.filename)
        except Exception as e:
            logging.error(f"Error al guardar: {e}")

if __name__ == "__main__":
    app = VulnerabilityDashboard()
    if app.fetch_data():
        app.create_dashboard()