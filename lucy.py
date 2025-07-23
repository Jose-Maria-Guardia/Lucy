import tkinter as tk
from tkinter import scrolledtext
import openai
import threading
import speech_recognition as sr
import subprocess
import time
from pywinauto import findwindows
from pywinauto import Desktop
import json
import os

# Ruta absoluta del directorio del script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

SYSTEM_PROMPT = (
    'Eres un asistente que convierte peticiones en lenguaje natural en instrucciones estructuradas para automatización de Windows.\n'
    'Devuelve SIEMPRE un JSON válido con una estructura determinada.\n'
    'Por ejemplo, si la consulta es "abre la calculadora" la estructura sera la siguiente:\n\n'
    '{\n  "acciones": [\n    { "tipo": "abrir_app", "parametros": {"nombre_app": "calculadora" } }\n  ]\n}'
    'Por ejemplo, si la consulta es "calculadora" la estructura sera la siguiente:\n\n'
    '{\n  "acciones": [\n    { "tipo": "foco_app", "parametros": {"nombre_ventana": "Calculadora" } }\n  ]\n}'

)

def cargar_alias_apps(path="alias_apps.json"):
    try:
        full_path = os.path.join(SCRIPT_DIR, path)
        with open(full_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error cargando alias_apps.json: {e}")
        return {}

ALIAS_APPS = cargar_alias_apps()

def recargar_alias():
    global ALIAS_APPS
    try:
        ALIAS_APPS = cargar_alias_apps()
        if not ALIAS_APPS:
            raise ValueError("El archivo alias_apps.json está vacío o no se pudo cargar.")
        salida.config(state='normal')
        salida.insert(tk.END, "[Alias recargados desde alias_apps.json]\n", 'user')
        salida.config(state='disabled')
        salida.see(tk.END)
    except Exception as e:
        salida.config(state='normal')
        salida.insert(tk.END, f"[Error al recargar alias: {e}]\n", 'user')
        salida.config(state='disabled')
        salida.see(tk.END)

# primer_mensaje = True

def get_openai_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("No se ha encontrado la variable de entorno OPENAI_API_KEY.")
    return openai.OpenAI(api_key=api_key)

def obtener_hijos(elemento):
    try:
        hijos = []
        for child in elemento.children():
            hijo_info = {
                "title": child.window_text(),
                "class_name": child.friendly_class_name(),
                "control_type": getattr(child.element_info, 'control_type', None),
                "automation_id": getattr(child.element_info, 'automation_id', None),
                "texts": child.texts()
            }
            hijos.append(hijo_info)
        return hijos
    except Exception as e:
        return [f"Error obteniendo hijos: {e}"]

def extraer_contexto_windows():
    try:
        desktop = Desktop(backend="uia")
        # Extraer barra de tareas y sus hijos
        taskbar = desktop.window(class_name='Shell_TrayWnd')
        barra_tareas = {
            'exists': taskbar.exists(),
            'texts': taskbar.texts(),
            'control_type': taskbar.friendly_class_name(),
            'hijos': obtener_hijos(taskbar)
        }
        # Extraer ventana activa de forma compatible
        from pywinauto import findwindows
        ventana_activa_hwnd = findwindows.find_windows(active_only=True)[0]
        ventana_activa = desktop.window(handle=ventana_activa_hwnd)
        activa = {
            'title': ventana_activa.window_text(),
            'class_name': ventana_activa.friendly_class_name(),
            'texts': ventana_activa.texts(),
            'hijos': obtener_hijos(ventana_activa)
        }
        contexto = {
            'barra_tareas': barra_tareas,
            'ventana_activa': activa
        }
        return contexto
    except Exception as e:
        return {'error': f'Error extrayendo contexto de Windows: {e}'}

# Función para consultar el modelo de OpenAI
def consultar_openai(pregunta):
    # global primer_mensaje
    client = get_openai_client()
    try:
        contexto = extraer_contexto_windows()
        contexto_str = json.dumps(contexto, ensure_ascii=False, indent=2)
        mensaje_usuario = (
            f"[Contexto de Windows adjunto]:\n{contexto_str}\n"  # <-- El contexto real del entorno
            f"[Petición]:\n{pregunta}"
        )
        mensajes_modelo = [
          {"role": "system", "content": SYSTEM_PROMPT},
          {"role": "user", "content": mensaje_usuario}
        ]

        respuesta = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=mensajes_modelo
        )
        return respuesta.choices[0].message.content.strip()
    except Exception as e:
        return f"Error al consultar OpenAI: {e}"

def transcribir_voz():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            salida.config(state='normal')
            salida.insert(tk.END, "Esperando voz...\n", 'user')
            salida.config(state='disabled')
            salida.see(tk.END)
            ventana.update()
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
            texto = recognizer.recognize_google(audio, language='es-ES')
            entrada.delete('1.0', tk.END)
            entrada.insert('1.0', texto)
        except Exception as e:
            salida.config(state='normal')
            salida.insert(tk.END, f"Error reconociendo voz: {e}\n", 'user')
            salida.config(state='disabled')
            salida.see(tk.END)
            ventana.update()

def boton_voz_thread():
    hilo = threading.Thread(target=transcribir_voz)
    hilo.start()

# Función que se ejecuta al presionar el botón
def enviar_consulta(event=None):
    pregunta = entrada.get("1.0", tk.END).strip()
    if pregunta:
        salida.config(state='normal')
        salida.insert(tk.END, f"Tú: {pregunta}\n", 'user')
        salida.config(state='disabled')
        salida.see(tk.END)
        ventana.update()
        respuesta = consultar_openai(pregunta)
        salida.config(state='normal')
        salida.insert(tk.END, f"IA: {respuesta}\n\n", 'ai')
        # ---- Intentar ejecutar acciones si hay JSON válido ----
        try:
            if "\"acciones\"" in respuesta:
                resultado_acciones = ejecutar_acciones(respuesta)
                salida.insert(tk.END, f"[Ejecución]: {resultado_acciones}\n\n", 'user')
        except Exception as e:
            salida.insert(tk.END, f"[Ejecución]: Error al ejecutar acciones: {e}\n\n", 'user')
        salida.config(state='disabled')
        salida.see(tk.END)
        entrada.delete("1.0", tk.END)

def ejecutar_acciones(json_respuesta):
    try:
        data = json.loads(json_respuesta)
        if "acciones" not in data:
            return "No se encontraron acciones para ejecutar."
        resultados = []
        for accion in data["acciones"]:
            tipo = accion.get("tipo")
            parametros = accion.get("parametros", {})
            nombre_app = parametros.get("nombre_app")
            nombre_ventana = parametros.get("nombre_ventana")

            if tipo == "abrir_app":
                if not nombre_app:
                    resultados.append("Falta el nombre_app en la acción abrir_app.")
                    continue
                # Traducir alias solo en abrir_app
                nombre_app_exec = ALIAS_APPS.get(nombre_app.lower(), nombre_app)
                # Buscar si ya está abierta
                app_abierta = False
                ventanas = findwindows.find_elements(title_re=".*", class_name=None)
                for v in ventanas:
                    try:
                        proceso = v.process
                        modulo = findwindows.get_process_module(proceso)
                        if nombre_app_exec.lower() in (modulo or '').lower():
                            app_abierta = True
                            resultados.append(f"{nombre_app_exec} ya estaba abierta. No se ha hecho nada.")
                            break
                    except Exception:
                        continue
                if not app_abierta:
                    try:
                        subprocess.Popen(nombre_app_exec)
                        resultados.append(f"{nombre_app_exec} no estaba abierta. Se ha lanzado la aplicación.")
                    except Exception as e:
                        resultados.append(f"Error al abrir {nombre_app_exec}: {e}")

            elif tipo == "foco_app":
                if not nombre_ventana:
                    resultados.append("Falta el nombre_ventana en la acción foco_app.")
                    continue
                encontrada = False
                ventanas = findwindows.find_elements(title_re=".*", class_name=None)
                for v in ventanas:
                    try:
                        win = Desktop(backend="uia").window(handle=v.handle)
                        # Comprobamos por título de ventana:
                        if nombre_ventana.lower() in win.window_text().lower():
                            win.set_focus()
                            encontrada = True
                            resultados.append(f"Se ha dado el foco a la ventana de {nombre_ventana}.")
                            break
                    except Exception:
                        continue
                if not encontrada:
                    resultados.append(f"No se encontró una ventana abierta de {nombre_ventana} para darle el foco.")

        return "\n".join(resultados) if resultados else "No se ejecutó ninguna acción."
    except Exception as e:
        return f"Error ejecutando acciones: {e}"

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Consultas IA OpenAI")
ventana.geometry('550x430')

entrada = tk.Text(ventana, width=60, height=4)
entrada.pack(pady=10)
entrada.bind('<Control-Return>', enviar_consulta)

frame_botones = tk.Frame(ventana)
frame_botones.pack(pady=5)

boton = tk.Button(frame_botones, text="Enviar consulta", command=enviar_consulta)
boton.pack(side='left', padx=5)

boton_voz = tk.Button(frame_botones, text="Hablar consulta 🎤", command=boton_voz_thread)
boton_voz.pack(side='left', padx=5)

boton_recargar = tk.Button(frame_botones, text="Recargar alias", command=recargar_alias)
boton_recargar.pack(side='left', padx=5)

salida = scrolledtext.ScrolledText(ventana, wrap=tk.WORD, state='disabled', width=60, height=17)
salida.tag_config('user', foreground='blue')
salida.tag_config('ai', foreground='green')
salida.pack(pady=10)

ventana.mainloop()
