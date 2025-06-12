# ANÁLISIS DE EMOCIONES EN DOCUMENTOS DE TEXTO (.docx)

Este proyecto realiza el análisis de sentimientos y emociones en documentos `.docx` escritos en español, utilizando técnicas de procesamiento de lenguaje natural (NLP) e inteligencia artificial. Está orientado a apoyar tareas de análisis textual en contextos educativos, psicológicos o de investigación.

---

## Funcionalidades

- Carga automática de archivos `.docx` desde una carpeta local
- Preprocesamiento del texto: división por frases, limpieza básica
- Clasificación emocional de cada frase (joy, sadness, anger, love, fear, surprise)
- Exportación de resultados en:
  - Word: frases con su emoción y nivel de confianza
  - Excel: datos + gráficos automáticos
-  Gráficos generados:
  - Gráfico circular (distribución de emociones)
  - Gráfico de líneas (evolución emocional por frase)
  - Heatmap (emociones por documento)
  - Gráfico de burbujas (frecuencia vs confianza)

---

## Tecnologías utilizadas

- `transformers` (Hugging Face)
- `pandas`
- `python-docx`
- `openpyxl` + `xlsxwriter`
- `matplotlib` / `seaborn`

Modelo usado: [`j-hartmann/emotion-english-distilroberta-base`]

---

## Estructura del proyecto
- analizar_emociones/
   - analizar_emociones.py # Script principal de análisis
   - requirements.txt # Dependencias
   - README.md # Este archivo
     - documentos/ # Archivos .docx a analizar
     - resultados/
         - Limpieza_emociones.docx
         - Limpieza_emociones.xlsx
     - informe/
         - Informe_Analisis_Emociones.docx
