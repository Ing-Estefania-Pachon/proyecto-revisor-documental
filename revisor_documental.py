import docx
import pandas as pd
import json
import os

# Skill basado en la metodología: Fase 2 - Elementos de Estilo [cite: 430]
def skill_revisor_estilo(numero_linea, texto):
    hallazgos = []
    # Regla 3.2. APA: Acrónimos - Verificar si falta el paréntesis [cite: 436, 437]
    if "SIATA" in texto and "(" not in texto: 
        hallazgos.append({
            "Línea": numero_linea,
            "Fase": "Elementos de Estilo",
            "Criterio": "Acrónimos",
            "Cumple": "No",
            "Recomendación": "Definir el acrónimo SIATA por extenso la primera vez[cite: 437]."
        })
    return hallazgos

def generar_reportes(nombre_archivo):
    # Verificamos si el archivo existe para evitar errores
    if not os.path.exists(nombre_archivo):
        print(f"❌ Error: No se encuentra el archivo '{nombre_archivo}' en la carpeta.")
        return

    print(f"🚀 Analizando: {nombre_archivo}...")
    doc = docx.Document(nombre_archivo)
    data_reporte = []
    
    # 1. Lectura e Indexación [cite: 60]
    for i, para in enumerate(doc.paragraphs):
        linea_texto = para.text.strip()
        if linea_texto:
            errores = skill_revisor_estilo(i + 1, linea_texto)
            if errores:
                data_reporte.extend(errores)
            else:
                data_reporte.append({
                    "Línea": i + 1,
                    "Fase": "Revisión General",
                    "Criterio": "Estilo",
                    "Cumple": "Sí",
                    "Recomendación": "N/A"
                })

    # 3. Generar EXCEL [cite: 381]
    df = pd.DataFrame(data_reporte)
    df.to_excel("Reporte_Cumplimiento.xlsx", index=False)
    
    # 4. Generar WORD [cite: 6, 369]
    reporte_doc = docx.Document()
    reporte_doc.add_heading('Reporte de Revisión de Estilo', 0)
    reporte_doc.add_paragraph('Recomendaciones basadas en la metodología institucional[cite: 369]:')
    
    for item in data_reporte:
        if item["Cumple"] == "No":
            p = reporte_doc.add_paragraph()
            p.add_run(f"Línea {item['Línea']}: ").bold = True
            p.add_run(item['Recomendación'])
            
    reporte_doc.save("Reporte_Recomendaciones.docx")
    print("✅ ¡Éxito! Se han generado: Reporte_Cumplimiento.xlsx y Reporte_Recomendaciones.docx")

# --- ESTA ES LA PARTE IMPORTANTE ---
if __name__ == "__main__":
    # Escribe AQUÍ el nombre exacto de tu archivo entre comillas
    generar_reportes("FICHA_CORDOBA.docx")