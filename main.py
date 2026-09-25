import openai

# Configura tu API Key aquí
client = openai.OpenAI(api_key="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")

def generar_informe_limpio():
    print("--- Creador de Informes (Sin Doble Título y Sin Errores) ---")
    
    instrucciones_estilo = (
        "Escribe como un estudiante de secundaria. "
        "REGLA DE ORO: No escribas ningún título ni encabezado al inicio de tu respuesta. Empieza directamente con el párrafo. "
        "Uso estricto de minúsculas: Todo el texto debe ir en minúsculas, EXCEPTO la primera letra de cada oración. "
        "Formato: No uses asteriscos (*) ni guiones (-) a menos que sea una lista de más de 3 elementos. "
        "Lenguaje: Natural y fluido. No hagas resúmenes ni conclusiones al final. "
        "EXTENSIÓN: El texto debe tener entre 110 y 135 palabras exactas. No te pases."
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

    print("\nGenerando documento final corregido...")

    documento_final = ""

    for i, punto in enumerate(puntos_usuario):
        print(f"[{i+1}/{cantidad_puntos}] Procesando: {punto}...")
        
        # Le pedimos a la IA que primero corrija el título para que no salga con errores
        # O lo corregimos directamente en el prompt
        response = client.chat.completions.create(
            model="gpt-4o", 
            messages=[
                {"role": "system", "content": instrucciones_estilo},
                {"role": "user", "content": f"Desarrolla el tema: {punto}. Recuerda no poner títulos y mantenerte bajo las 135 palabras."}
            ],
            temperature=0.5,
            frequency_penalty=0.9 
        )
        
        texto_generado = response.choices[0].message.content
        
        # Limpieza de asteriscos por seguridad
        texto_generado = texto_generado.replace("*", "")
        
        # Corregimos el título para que no salga con errores (como 'QUES' o 'VENEZEULA')
        # Pasamos el punto a mayúsculas y corregimos errores comunes manualmente
        titulo_corregido = punto.upper().replace("QUES", "QUE").replace("VENEZEULA", "VENEZUELA").replace("INFROME", "INFORME")
        
        # Unimos: Título corregido + Texto generado (que ya no trae título interno)
        documento_final += f"{titulo_corregido}\n\n{texto_generado}\n\n"

    nombre_archivo = "Informe_Final_Soberania.txt"
    with open(nombre_archivo, "w", encoding="utf-8") as archivo:
        archivo.write(documento_final)

    print(f"\n¡Listo! El archivo '{nombre_archivo}' no tiene títulos duplicados ni errores.")

if __name__ == "__main__":
    generar_informe_limpio()