import os
import openai
import base64
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DIRECTA ---
# Pega tu clave de OpenAI entre las comillas para que el script funcione sin preguntar nada.
api_key = "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"

client = openai.OpenAI(api_key=api_key)

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def extraer_temas_de_imagen(ruta_imagen):
    print(f"\n--- Analizando imagen detalladamente: {ruta_imagen} ---")
    base64_image = encode_image(ruta_imagen)
    
    prompt_vision = (
        "Analiza esta imagen escolar. "
        "1. Identifica el TEMA GLOBAL de la asignación. "
        "2. Identifica todos los puntos con viñetas o asteriscos. "
        "REGLA CRÍTICA: Devuelve cada uno de los elementos como un tema individual separado por comas. "
        "No resumas ni mezcles los conceptos de diferentes personajes. "
        "Devuelve solo los nombres de los temas separados por comas, sin etiquetas."
    )

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": [
            {"type": "text", "text": prompt_vision},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
        ]}]
    )
    
    contenido = response.choices[0].message.content
    for etiqueta in ["Tema principal:", "Tema global:", "Puntos:", "Subpuntos:"]:
        contenido = contenido.replace(etiqueta, "")
        
    return [tema.strip() for tema in contenido.split(',') if tema.strip()]

def procesar_texto_manual():
    print("\n--- Ingreso de asignación por texto ---")
    print("Escribe o pega los temas separados por comas (ejemplo: Gobierno de Juan Vicente Gomez, Gobierno de Perez Jimenez):")
    entrada = input("> ")
    return [tema.strip() for tema in entrada.split(',') if tema.strip()]

def corregir_capitales_y_ortografia(texto):
    """
    Filtro estricto que asegura mayúsculas en nombres propios, apellidos y lugares de la historia venezolana.
    """
    reemplazos = {
        r"\bgomez\b": "Gómez",
        r"\bperez\b": "Pérez",
        r"\bjimenez\b": "Jiménez",
        r"\bjuan\b": "Juan",
        r"\bvicente\b": "Vicente",
        r"\bmarcos\b": "Marcos",
        r"\bvenezuela\b": "Venezuela",
        r"\bcaracas\b": "Caracas",
    }
    
    for patron, reemplazo in reemplazos.items():
        # Corregido: Se eliminó 'text=' para evitar el TypeError
        texto = re.sub(patron, reemplazo, texto, flags=re.IGNORECASE)
    
    texto = re.sub(r'(^[a-z]|(?<=\.\s)[a-z])', lambda m: m.group(1).upper(), texto)
    return texto

def generar_informe_final():
    print("¿Cómo deseas ingresar la asignación?")
    print("1. Mediante una Imagen")
    print("2. Mediante Texto Manual")
    opcion = input("Selecciona una opción (1 o 2): ").strip()

    if opcion == "1":
        ruta = input("Ingresa la ruta de la imagen: ").strip('"')
        temas_extraidos = extraer_temas_de_imagen(ruta)
    elif opcion == "2":
        temas_extraidos = procesar_texto_manual()
    else:
        print("Opción no válida. Saliendo del programa.")
        return
    
    palabras_prohibidas = ["soberanía", "4to año", "lunes", "investigación", "20/04"]
    temas_finales = [t for t in temas_extraidos if not any(p in t.lower() for p in palabras_prohibidas)]

    if not temas_finales:
        print("No se detectaron temas válidos.")
        return

    doc = Document()
    print(f"\nTemas detectados para redactar: {len(temas_finales)}")

    for i, tema in enumerate(temas_finales):
        numero = i + 1
        
        es_tema_importante = any(palabra in tema.lower() for palabra in ["gobierno", "transicion", "transición"])
        
        if es_tema_importante:
            rango_palabras = "entre 160 y 180 palabras (un desarrollo extenso, profundo y detallado)"
            print(f"[{numero}/{len(temas_finales)}] Redactando Tema Principal (Extendido): {tema}...")
        else:
            rango_palabras = "entre 115 y 125 palabras"
            print(f"[{numero}/{len(temas_finales)}] Redactando Subtema: {tema}...")
        
        t_clean = tema.replace("-", "-").replace("•", "").strip()
        t_clean = corregir_capitales_y_ortografia(t_clean.strip('¿?'))
        if len(t_clean) > 0:
            t_clean = t_clean[0].upper() + t_clean[1:]
        
        t_low = t_clean.lower()
        es_pregunta = t_low.startswith(("que", "como", "cual", "por que", "por qué", "quien", "donde"))
        titulo_final = f"{numero}. ¿{t_clean}?" if es_pregunta else f"{numero}. {t_clean}"

        instrucciones_redaccion = (
            f"Eres un historiador experto en historia contemporánea de Venezuela. Desarrolla el tema asignado con absoluto rigor académico. "
            f"REGLA CRÍTICA DE CONTEXTO: Enfócate ÚNICAMENTE en los personajes e hitos mencionados en el título de la tarea. No mezcles administraciones si el título no lo pide. "
            f"REGLA CRÍTICA DE EXTENSIÓN: El texto completo debe tener {rango_palabras}. "
            f"No incluyas títulos en tu respuesta. Empieza directo. Prohibido usar punto y coma (;). "
            f"MINÚSCULAS Y MAYÚSCULAS: Todo el texto regular en minúsculas, EXCEPTO la primera letra de cada oración y la primera letra de nombres propios de personas y lugares. "
            f"LISTAS: Usa •. Formato 'Componente: descripción breve' (máximo 20 palabras por punto)."
        )

        formato_lista = " DEBES incluir una lista formal estructurada con viñetas 'Componente: descripción'." if i % 2 == 0 else " No uses listas, redacta completamente en párrafos continuos."
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": instrucciones_redaccion},
                {"role": "user", "content": f"Desarrolle exclusivamente el punto: '{tema}'.{formato_lista}"}
            ],
            temperature=0.3
        )
        
        texto_generado = response.choices[0].message.content.replace("*", "").replace(";", ".")
        texto_generado = corregir_capitales_y_ortografia(texto_generado)

        # --- FORMATO WORD ---
        h = doc.add_paragraph()
        run_h = h.add_run(titulo_final)
        run_h.bold = True
        run_h.font.name = 'Canva Sans'
        run_h.font.size = Pt(15)

        lineas = texto_generado.split('\n')
        for linea in lineas:
            linea = linea.strip()
            if not linea: continue
            
            if not linea.startswith("•") and len(linea.split()) > 65:
                puntos = linea.split('. ')
                mitad = len(puntos) // 2
                bloques = [". ".join(puntos[:mitad]) + ".", ". ".join(puntos[mitad:])] if mitad > 0 else [linea]
            else:
                bloques = [linea]

            for bloque in bloques:
                p = doc.add_paragraph()
                if bloque.strip().startswith("•"):
                    p.add_run("• ").bold = True
                    cont = bloque.strip().lstrip("• ").strip()
                    if ":" in cont:
                        sub, desc = cont.split(":", 1)
                        run_sub = p.add_run(f"{sub.strip()}:")
                        run_sub.bold = True
                        p.add_run(desc)
                    else:
                        p.add_run(cont)
                    p.paragraph_format.left_indent = Pt(24)
                else:
                    p.add_run(bloque.strip())
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                for run in p.runs:
                    run.font.name = 'Canva Sans'
                    run.font.size = Pt(13)

        doc.add_paragraph()

    doc.save("Informe_Historico_Personalizado.docx")
    print(f"\n¡Listo! Se procesaron {len(temas_finales)} temas individuales sin cruzar información.")

if __name__ == "__main__":
    generar_informe_final()