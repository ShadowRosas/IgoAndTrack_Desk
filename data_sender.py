import pandas as pd
import datetime
import requests
import pytz

def enviar_datos(archivo_excel, update_message, stop_event):
    url = 'https://teach2.goandsee.co/api/v3/oee'
    
    try:
        # Obtener todos los nombres de las hojas del archivo
        xls = pd.ExcelFile(archivo_excel)
        hojas = xls.sheet_names
        
        for nombre_hoja in hojas:
            if stop_event.is_set():
                update_message("Proceso de envío de datos detenido.")
                return

            try:
                df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
                if df.shape[1] < 8:
                    raise ValueError(f"La hoja {nombre_hoja} no tiene suficientes columnas.")
                
                qty_column = pd.to_numeric(df.iloc[:, 5], errors='coerce')
                tijuana_tz = pytz.timezone('America/Tijuana')
                hour_column = pd.to_datetime(df.iloc[:, 7], errors='coerce', format='%H:%M:%S')
                hour_column = hour_column.dt.tz_localize(tijuana_tz, ambiguous='NaT', nonexistent='shift_forward')
                
                pulsos_simplificados = []
                for idx, (qty, hour) in enumerate(zip(qty_column, hour_column)):
                    if stop_event.is_set():
                        update_message("Proceso de envío de datos detenido.")
                        return

                    if pd.notna(qty) and qty > 0:
                        initial_time = hour.replace(minute=0, second=0, microsecond=0)
                        total_minutes = hour.minute + hour.second / 60
                        interval_minutes = total_minutes / (qty - 1) if qty > 1 else total_minutes
                        for i in range(int(qty)):
                            next_time = initial_time + datetime.timedelta(minutes=interval_minutes * i)
                            pulsos_simplificados.append({
                                'datetime': next_time.strftime('%H:%M:%S'),
                                'date': int(next_time.timestamp() * 1000),
                                'eth_mac': "1c:2c:3c",  
                                'sensor': 10,
                            })
                
                for pulse_data in pulsos_simplificados:
                    if stop_event.is_set():
                        update_message("Proceso de envío de datos detenido.")
                        return

                    message = f"Pulso enviado desde la hoja {nombre_hoja}:\n{format_pulse_data(pulse_data)}"
                    print(message)
                    update_message(message)
                    response = requests.post(url, json=pulse_data)
                    if response.status_code != 200:
                        error_message = f"Error al enviar datos al servidor desde la hoja {nombre_hoja}: {response.status_code}"
                        print(error_message)
                        update_message(error_message)
                
                success_message = f"Datos de la hoja {nombre_hoja} han sido enviados con éxito."
                print(success_message)
                update_message(success_message)
            
            except Exception as e:
                error_message = f"Error en la hoja {nombre_hoja}: {e}"
                print(error_message)
                update_message(error_message)
                continue
        
        success_message = "Todos los datos han sido enviados con éxito desde todas las hojas."
        print(success_message)
        update_message(success_message)
        
    except Exception as e:
        error_message = f"Error al procesar el archivo Excel: {e}"
        print(error_message)
        update_message(error_message)

def format_pulse_data(pulse_data):
    formatted_data = (
        f"Fecha y hora: {pulse_data['datetime']}\n"
        f"Fecha (timestamp): {pulse_data['date']}\n"
        f"MAC Ethernet: {pulse_data['eth_mac']}\n"
        f"Sensor: {pulse_data['sensor']}\n"
    )
    return formatted_data
