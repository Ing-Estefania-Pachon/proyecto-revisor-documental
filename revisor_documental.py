import os
import docx
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv
import json

# Cargar configuración segura
load_dotenv()
genai.configure(api_key=os.getenv("ARES_API_KEY"))
model = genai.GenerativeModel('gemini-1.5-flash')

def llamar_ia(prompt):
    response = model.generate_content(prompt)
    # Limpiar respuesta para obtener JSON puro
    res_text = response.text.replace('```json', '').replace('```', '').strip()
    return json.loads(res_text)

def generar_reporte_ares(nombre_archivo):
    doc = docx.Document(nombre_archivo)
    # Fase 1: Metadatos (Primeros párrafos)
    texto_fase1 = "\n".join([p.text for p in doc.paragraphs[:15]])
    
    # Aquí puedes cambiar el prompt según la fase que necesites correr
    print("🤖 Analizando Fase 1...")
    resultado = llamar_ia(f"Actúa como auditor UNGRD. Analiza Fase 1: {texto_fase1}")
    
    # Guardar Excel
    pd.DataFrame(resultado["resultados"]).to_excel("Reporte_Cumplimiento.xlsx", index=False)
    
    # Guardar Word
    reporte = docx.Document()
    reporte.add_heading('ARES - Reporte de Revisión', 0)
    for res in resultado["resultados"]:
        reporte.add_paragraph(f"{res['criterio']}: {res['estado']} - {res['recomendacion']}")
    reporte.save("Reporte_Recomendaciones.docx")
    print("✅ Reportes generados exitosamente.")

if __name__ == "__main__":
    generar_reporte_ares("FICHA_CORDOBA.docx")