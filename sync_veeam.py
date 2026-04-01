import win32com.client
import psycopg2
import sys

# Configuración para que la consola de Windows no de errores con caracteres especiales
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Configuración de tu base de datos PostgreSQL
DB_CONFIG = {
    "dbname": "veeam_monitor",
    "user": "postgres",
    "password": "CONTRASEÑA BBDD",
    "host": "127.0.0.1"
}

def sincronizar_outlook():
    try:
        # 1. Conexión a Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Seleccionamos la cuenta específica
        cuenta = next((s for s in outlook.Stores if "alarmas@laberit.com" in s.DisplayName.lower()), None)
        if not cuenta:
            print("❌ No se encontró la cuenta alarmas@laberit.com en Outlook.")
            return

        root = cuenta.GetRootFolder()
        inbox = next((f for f in root.Folders if f.Name.lower() in ["bandeja de entrada", "inbox"]), None)
        
        if not inbox:
            print("❌ No se encontró la Bandeja de Entrada.")
            return

        # 2. Conexión a PostgreSQL
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        print(f"🔄 Sincronizando subcarpetas de {inbox.Name} (extrayendo logs)...")

        # 3. Recorrer cada subcarpeta (cada Cliente)
        for subfolder in inbox.Folders:
            nombre_cliente = subfolder.Name
            mensajes = subfolder.Items
            
            if mensajes.Count == 0:
                continue

            mensajes.Sort("[ReceivedTime]", True)
            
            # Límite de 50 para que el Dashboard sea ágil
            limite = min(51, mensajes.Count + 1)
            nuevos_o_actualizados = 0

            for i in range(1, limite):
                try:
                    msg = mensajes.Item(i)
                    asunto = msg.Subject
                    fecha = msg.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
                    
                    # --- NUEVO: Extraer el cuerpo del correo ---
                    cuerpo_log = msg.Body 
                    
                    # Clasificación del estado
                    as_up = asunto.upper()
                    if any(x in as_up for x in ["FAILED", "ERROR"]):
                        estado = "Failed"
                    elif "WARNING" in as_up:
                        estado = "Warning"
                    else:
                        estado = "Success"

                    # --- MODIFICADO: Insertar incluyendo log_cuerpo ---
                    # Si ya existe (CONFLICT), actualizamos el log_cuerpo por si estaba vacío
                    cur.execute("""
                        INSERT INTO backups (cliente, job_name, status, fecha, log_cuerpo)
                        VALUES (%s, %s, %s, %s, %s)
                        ON CONFLICT (cliente, job_name, fecha) 
                        DO UPDATE SET log_cuerpo = EXCLUDED.log_cuerpo 
                        WHERE backups.log_cuerpo IS NULL;
                    """, (nombre_cliente, asunto, estado, fecha, cuerpo_log))
                    
                    if cur.rowcount > 0:
                        nuevos_o_actualizados += 1
                        
                except Exception as e:
                    continue 

            if nuevos_o_actualizados > 0:
                print(f"✅ {nombre_cliente}: {nuevos_o_actualizados} registros procesados.")
            
            conn.commit()

        cur.close()
        conn.close()
        print("\n✨ Sincronización finalizada con éxito con logs completos.")

    except Exception as e:
        print(f"❌ Error crítico durante la sincronización: {e}")

if __name__ == "__main__":
    sincronizar_outlook()