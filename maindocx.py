import openai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configura tu API Key aquí
client = openai.OpenAI(api_key="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")

def generar_word_profesional_canva():
    print("--- Creador de Informes (Control Estricto de Palabras) ---")
    
    instrucciones_estilo = (
        "Adopta un tono académico formal. "
        "REGLA CRÍTICA: El texto completo (incluyendo listas) DEBE tener entre 110 y 125 palabras. "
        "Si te pasas de 130 palabras, el trabajo será rechazado. Sé muy sintético. "
        "No incluyas títulos. Empieza directo. "
        "SIGNOS: No uses punto y coma (;). "
        "MINÚSCULAS: Todo en minúsculas, EXCEPTO la primera letra de cada oración. "
        "LISTAS: Usa el símbolo •. El formato debe ser 'Componente: descripción breve'. "
        "La descripción de cada punto de la lista no debe exceder las 20 palabras."
    )

    try:
        cantidad_puntos = int(input("¿Cuántos puntos tiene tu trabajo?: "))
    except ValueError:
        print("Número no válido.")
        return

    puntos_usuario = []
    for i in range(cantidad_puntos):
        punto = input(f"Escribe el tema del punto {i+1}: ")
        puntos_usuario.append(punto)

    doc = Document()
    
    print("\nRedactando con control de extensión...")

    for i, punto in enumerate(puntos_usuario):
        numero_punto = i + 1
        print(f"[{numero_punto}/{cantidad_puntos}] Procesando: {punto}...")
        
        t_limpio = punto.lower().replace("truismo", "turismo").replace("venezeula", "venezuela").replace("ques", "que")
        
        if t_limpio.startswith(("que", "como", "cual", "donde", "por que", "quien")):
            texto_titulo = f"¿{t_limpio.capitalize()}?"
        else:
            texto_titulo = t_limpio.capitalize()
            
        titulo_final = f"{numero_punto}. {texto_titulo}"

        # Instrucción de lista solo si es necesario, pero reforzando la brevedad
        formato_lista = " Incluye una lista de máximo 3 componentes con descripciones muy cortas." if i % 2 == 0 else " No uses listas."
        
        response = client.chat.completions.create(
            model="gpt-4o", 
            messages=[
                {"role": "system", "content": instrucciones_estilo},
                {"role": "user", "content": f"Tema: {texto_titulo}.{formato_lista} Recuerda: Máximo 125 palabras en total."}
            ],
            temperature=0.4, # Menor temperatura = más directo y menos rollero
            presence_penalty=0.5 # Evita que se extienda repitiendo conceptos
        )
        
        texto_completo = response.choices[0].message.content.replace("*", "").replace(";", ".")
        
        # --- AGREGAR TÍTULO (15pt) ---
        h = doc.add_paragraph()
        run_h = h.add_run(titulo_final)
        run_h.bold = True
        run_h.font.name = 'Canva Sans'
        run_h.font.size = Pt(15)

        # --- PROCESAR CUERPO ---
        lineas = texto_completo.split('\n')
        for linea in lineas:
            linea = linea.strip()
            if not linea: continue
            
            # División de párrafos si exceden 60 palabras
            if not linea.startswith("•") and len(linea.split()) > 60:
                frases = linea.split('. ')
                punto_medio = len(frases) // 2
                partes_a_escribir = [". ".join(frases[:punto_medio]) + ".", ". ".join(frases[punto_medio:])]
            else:
                partes_a_escribir = [linea]

            for parte in partes_a_escribir:
                if not parte.strip(): continue
                p = doc.add_paragraph()
                
                if parte.startswith("•"):
                    contenido_linea = parte.lstrip("• ").strip()
                    run_v = p.add_run("• ")
                    run_v.bold = True
                    run_v.font.name = 'Canva Sans'
                    run_v.font.size = Pt(13)

                    if ":" in contenido_linea:
                        sub, desc = contenido_linea.split(":", 1)
                        run_sub = p.add_run(f"{sub.strip()}:")
                        run_sub.bold = True
                        run_sub.font.name = 'Canva Sans'
                        run_sub.font.size = Pt(13)
                        run_desc = p.add_run(desc)
                        run_desc.font.name = 'Canva Sans'
                        run_desc.font.size = Pt(13)
                    else:
                        run_txt = p.add_run(contenido_linea)
                        run_txt.font.name = 'Canva Sans'
                        run_txt.font.size = Pt(13)
                    p.paragraph_format.left_indent = Pt(24)
                else:
                    run_p = p.add_run(parte.strip())
                    run_p.font.name = 'Canva Sans'
                    run_p.font.size = Pt(13)
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        doc.add_paragraph()

    doc.save("Informe_Soberania_Final_Enumerado.docx")
    print("\n¡Archivo generado! Ahora el texto es mucho más corto y cabe en Canva.")

if __name__ == "__main__":
    generar_word_profesional_canva()