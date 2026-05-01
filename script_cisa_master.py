import pandas as pd
import requests
import matplotlib.pyplot as plt
import logging
import os
from datetime import datetime, timedelta
from fpdf import FPDF

# Configuración de logs profesional para auditoría de sistemas[cite: 3]
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class VulnerabilityAuditor:
    """
    Sistema Automatizado de Inteligencia de Amenazas (CISA KEV).
    Diseñado para integración en flujos de trabajo de ciberseguridad corporativa.[cite: 1]
    """
    def __init__(self):
        self.url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
        self.df = None
        self.report_date = datetime.now().strftime("%Y-%m-%d")

    def fetch_intelligence(self, days_back=365):
        """Descarga y procesa datos del feed oficial sin intervención manual.[cite: 1]"""
        try:
            logging.info("Sincronizando con el catálogo CISA KEV...")
            response = requests.get(self.url, timeout=15)
            response.raise_for_status()
            
            # Carga de datos y tipado de fechas
            self.df = pd.DataFrame(response.json()['vulnerabilities'])
            self.df['dateAdded'] = pd.to_datetime(self.df['dateAdded'])
            
            # Filtro temporal automático
            cutoff = datetime.now() - timedelta(days=days_back)
            self.df = self.df[self.df['dateAdded'] >= cutoff]
            
            logging.info(f"Análisis completado: {len(self.df)} vulnerabilidades críticas detectadas.")
            return True
        except Exception as e:
            logging.error(f"Error en la recolección de datos: {e}")
            return False

    def export_and_open(self):
        """Genera reportes técnicos y abre el archivo Excel inmediatamente.[cite: 1, 3]"""
        excel_file = f"Auditoria_KEV_{self.report_date}.xlsx"
        pdf_file = f"Reporte_Ejecutivo_{self.report_date}.pdf"

        # Exportación a Excel (Requiere openpyxl)[cite: 1]
        export_df = self.df.copy()
        export_df['dateAdded'] = export_df['dateAdded'].dt.date
        export_df.to_excel(excel_file, index=False, sheet_name='Analisis_KEV')
        
        # Generación de visualización estadística[cite: 1, 3]
        plt.figure(figsize=(10, 6))
        self.df['vendorProject'].value_counts().head(10).plot(kind='barh', color='#0d2a47')
        plt.title(f"Principales Amenazas Detectadas ({self.report_date})")
        plt.gca().invert_yaxis()
        plt.tight_layout()
        plt.savefig("chart_temp.png")

        # Generación de PDF ejecutivo[cite: 1]
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "REPORTE DE INTELIGENCIA DE AMENAZAS", 0, 1, 'C')
        pdf.image("chart_temp.png", x=10, y=30, w=190)
        pdf.output(pdf_file)

        logging.info(f"Archivos generados con éxito: {excel_file} y {pdf_file}")
        
        # Apertura automática en Windows[cite: 3]
        os.startfile(excel_file)

if __name__ == "__main__":
    auditor = VulnerabilityAuditor()
    if auditor.fetch_intelligence(days_back=365):
        auditor.export_and_open()