import time
from datetime import datetime, timedelta
import winsound

# ===== CONFIGURACIÓN =====
hora_alarma = "03:20"      # <-- FORMATO HH:MM (24 horas)
duracion_alarma = 5 * 60  # 5 minutos
intervalo = 1             # beep cada 1 segundo
# =========================

# Convertir a datetime
ahora = datetime.now()
hora_obj = datetime.strptime(hora_alarma, "%H:%M").time()
tiempo_alarma = datetime.combine(ahora.date(), hora_obj)

# Si la hora ya pasó hoy, programar para mañana
if tiempo_alarma <= ahora:
    tiempo_alarma += timedelta(days=1)

print(f"⏰ Alarma programada para: {tiempo_alarma.strftime('%Y-%m-%d %H:%M:%S')}")

# Esperar hasta la alarma
while datetime.now() < tiempo_alarma:
    time.sleep(1)

print("🚨 ¡ALARMA ACTIVADA!")

inicio = time.time()
while time.time() - inicio < duracion_alarma:
    winsound.Beep(1000, 500)
    time.sleep(intervalo)
