import tkinter as tk
from tkinter import filedialog
import time
from datetime import datetime
import threading
from data_sender import enviar_datos

class TimerApp:
    def __init__(self, root):
        self.root = root
        self.start_time = None  # Inicialmente, no se ha iniciado el contador
        self.running = False
        self.stop_event = threading.Event()
        self.selected_file = None

        self.root.title("iGoAndTrack")
        self.root.geometry("600x600")  # Tamaño de la ventana aumentado
        self.root.configure(bg="#282c34")  # Fondo de color oscuro

        # Fuente personalizada
        self.title_font = ("Helvetica", 18, "bold")
        self.label_font = ("Helvetica", 14)
        self.button_font = ("Helvetica", 12)

        # Título
        self.title_label = tk.Label(root, text="iGoAndTrack", font=self.title_font, bg="#61afef", fg="#282c34")
        self.title_label.pack(pady=20, padx=20, fill=tk.X)

        # Contenido en un marco
        self.content_frame = tk.Frame(root, bg="#282c34")
        self.content_frame.pack(pady=10)

        self.timer_label = tk.Label(self.content_frame, text="", font=self.label_font, bg="#282c34", fg="#98c379")
        self.timer_label.pack(pady=10)

        self.time_label = tk.Label(self.content_frame, text="", font=self.label_font, bg="#282c34", fg="#e06c75")
        self.time_label.pack(pady=10)

        self.message_text = tk.Text(self.content_frame, height=7, width=60, font=self.label_font, bg="#282c34", fg="#d19a66")
        self.message_text.pack(pady=10)

        self.select_file_button = tk.Button(root, text="Seleccionar Archivo Excel", command=self.select_file, font=self.button_font, bg="#61afef", fg="#282c34", activebackground="#98c379")
        self.select_file_button.pack(pady=10)

        self.send_data_button = tk.Button(root, text="Enviar Datos", command=self.run_enviar_datos, font=self.button_font, bg="#61afef", fg="#282c34", activebackground="#98c379")
        self.send_data_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Detener", command=self.stop_timer, font=self.button_font, bg="#e06c75", fg="#282c34", activebackground="#98c379")
        self.stop_button.pack(pady=10)

        self.update_time_label()
        
    def select_file(self):
        self.selected_file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm")],
            title="Seleccionar Archivo Excel"
        )
        if self.selected_file:
            self.update_message(f"Archivo seleccionado: {self.selected_file}")

    def update_time_label(self):
        if self.start_time is not None and self.running:
            elapsed_time = time.time() - self.start_time
            formatted_elapsed_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
            self.timer_label.config(text=f"Tiempo transcurrido: {formatted_elapsed_time}")

        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.config(text=f"Hora actual: {current_time}")

        self.root.after(1000, self.update_time_label)  # Actualiza cada segundo

    def run_enviar_datos(self):
        if self.selected_file is None:
            self.update_message("Por favor, seleccione un archivo Excel primero.")
            return
        if not self.running:
            self.start_time = time.time()  # Inicia el contador de tiempo
            self.running = True
            self.stop_event.clear()  # Asegúrate de que el evento de detener esté limpio
        threading.Thread(target=self.enviar_datos_con_mensaje).start()

    def stop_timer(self):
        self.running = False  # Detiene el contador de tiempo
        self.stop_event.set()  # Señala que el envío de datos debe detenerse

    def enviar_datos_con_mensaje(self):
        self.update_message("Iniciando el proceso de envío de datos...")
        enviar_datos(self.selected_file, self.update_message, self.stop_event)
        self.update_message("Proceso de envío de datos completado.")

    def update_message(self, message):
        self.message_text.insert(tk.END, message + "\n")
        self.message_text.see(tk.END)
