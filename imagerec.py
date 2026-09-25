import openai
import base64

# Configura tu API Key
client = openai.OpenAI(api_key="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")

def encode_image(image_path):
    """Convierte la imagen a base64 para enviarla a la API."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def extraer_temas_coherentes(ruta_imagen):
    print(f"--- Analizando imagen: {ruta_imagen} ---")
    
    # Codificar la imagen
    base64_image = encode_image(ruta_imagen)
    
    # Prompt específico para ignorar dibujos y solo extraer temas
    prompt_vision = (
        "Analiza esta imagen de una asignatura escolar. "
        "INSTRUCCIONES: "
        "1. Ignora cualquier dibujo, garabato, margen o texto irrelevante. "
        "2. Identifica los títulos o puntos clave que servirían para un informe. "
        "3. Corrige la ortografía de los temas detectados (ej. 'truismo' -> 'turismo'). "
        "4. Devuelve los temas en una lista simple separada por comas (ejemplo: Tema 1, Tema 2, Tema 3)."
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_vision},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        },
                    ],
                }
            ],
            max_tokens=300
        )

        # Extraer el texto de la respuesta
        contenido = response.choices[0].message.content
        
        # Convertir el texto en una lista de Python para que sea útil
        lista_temas = [tema.strip() for tema in contenido.split(',')]
        
        return lista_temas

    except Exception as e:
        return f"Error al procesar la imagen: {e}"

# --- PRUEBA DEL MÓDULO ---
if __name__ == "__main__":
    # Aquí pones la ruta de tu foto
    ruta = input("Ingresa la ruta de la imagen (o arrástrala): ").strip('"')
    
    puntos_importantes = extraer_temas_coherentes(ruta)
    
    print("\n--- Puntos importantes detectados ---")
    if isinstance(puntos_importantes, list):
        for i, tema in enumerate(puntos_importantes, 1):
            print(f"{i}. {tema}")
    else:
        print(puntos_importantes)