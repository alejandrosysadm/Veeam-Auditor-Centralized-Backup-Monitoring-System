import win32com.client
import psycopg2
import sys

# Configuración de salida segura para evitar errores con caracteres especiales en Windows
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DB_CONFIG = {
    "dbname": "veeam_monitor",
    "user": "postgres",
    "password": "CONTRASEÑA BBDD",
    "host": "127.0.0.1"
}

def carga_masiva():
    try:
        print("🚀 Iniciando CARGA HISTÓRICA con LOGS (Límite: 500 por carpeta)...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Seleccionamos la cuenta específica
        cuenta = next((s for s in outlook.Stores if "CUENTA A MONITORIZAR" in s.DisplayName.lower()), None)
        if not cuenta:
            print("❌ No se encontró la cuenta CUENTA A MONITORIZAR en Outlook.")
            return

        root = cuenta.GetRootFolder()
        inbox = next((f for f in root.Folders if f.Name.lower() in ["bandeja de entrada", "inbox"]), None)
        
        if not inbox:
            print("❌ No se encontró la Bandeja de Entrada.")
            return
        
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        for subfolder in inbox.Folders:
            nombre_cliente = subfolder.Name
            mensajes = subfolder.Items
            if mensajes.Count == 0: continue
            
            # Ordenar por fecha para asegurar que traemos los más recientes primero
            mensajes.Sort("[ReceivedTime]", True)
            
            # LEEMOS LOS ÚLTIMOS 500 CORREOS (Ajustado según el límite definido en el script original)
            limite = min(501, mensajes.Count + 1)
            nuevos = 0
            
            for i in range(1, limite):
                try:
                    msg = mensajes.Item(i)
                    asunto = msg.Subject
                    fecha = msg.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
                    
                    # --- NUEVO: Extraer el cuerpo del correo (Log) ---
                    cuerpo_log = msg.Body
                    
                    # Clasificación del estado
                    as_up = asunto.upper()
                    if any(x in as_up for x in ["FAILED", "ERROR"]): 
                        est = "Failed"
                    elif "WARNING" in as_up: 
                        est = "Warning"
                    else: 
                        est = "Success"

                    # --- MODIFICADO: Incluimos log_cuerpo en el INSERT ---
                    cur.execute("""
                        INSERT INTO backups (cliente, job_name, status, fecha, log_cuerpo)
                        VALUES (%s, %s, %s, %s, %s)
                        ON CONFLICT (cliente, job_name, fecha) 
                        DO UPDATE SET log_cuerpo = EXCLUDED.log_cuerpo 
                        WHERE backups.log_cuerpo IS NULL;
                    """, (nombre_cliente, asunto, est, fecha, cuerpo_log))
                    
                    if cur.rowcount > 0: 
                        nuevos += 1
                except Exception:
                    continue
            
            if nuevos > 0:
                print(f"📦 {nombre_cliente}: {nuevos} registros añadidos/actualizados con logs.")
            conn.commit()

        cur.close()
        conn.close()
        print("\n✅ HISTÓRICO COMPLETADO. Ya se han cargado los estados y cuerpos de los mensajes.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    carga_masiva()
