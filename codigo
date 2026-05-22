import pandas as pd
import requests
import datetime
import os
import time

def buscar_hospitales_excel():
    # --- 1. CONFIGURACIÓN ---
    API_KEY = os.getenv('MP_TICKET') 
    
    # Filtros enfocados puramente en salud e infraestructura
    FILTROS_HOSPITAL = ["hospital", "cesfam", "cecosf", "clinica", "posta", "salud"]

    proyectos_encontrados = []
    ids_procesados = set()
    hoy = datetime.datetime.now()

    print("Iniciando búsqueda de proyectos hospitalarios en Mercado Público...")

    try:
        # --- 2. CONSULTA A MERCADO PÚBLICO (Tu misma lógica) ---
        for i in range(3): 
            fecha_str = (hoy - datetime.timedelta(days=i)).strftime("%d%m%Y")
            url = f"https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json?fecha={fecha_str}&ticket={API_KEY}"

            for intento in range(2): 
                try:
                    print(f"Consultando {fecha_str} (Intento {intento+1})...")
                    response = requests.get(url, timeout=60)
                    
                    if response.status_code == 200:
                        data = response.json()
                        
                        if data and "Listado" in data:
                            if data["Listado"] is not None:
                                for lic in data["Listado"]:
                                    codigo = lic.get("CodigoExterno")
                                    if not codigo or codigo in ids_procesados: continue
                                    
                                    nombre = str(lic.get("Nombre", "")).lower()
                                    
                                    # Filtramos si el nombre tiene términos de salud
                                    if any(f in nombre for f in FILTROS_HOSPITAL):
                                        organismo = lic.get('OrganismoExterno', 'No especificado')
                                        
                                        # Armamos la fila del Excel con las columnas exactas de tu matriz
                                        proyectos_encontrados.append({
                                            'Nombre del Proyecto / Hospital': lic.get('Nombre'),
                                            'Región / Ciudad': organismo,
                                            'Constructora a Cargo': 'Por revisar en bases', 
                                            'Estado General del Proyecto': 'Licitación Publicada',
                                            '% de Avance Físico General': 'N/A',
                                            'Etapa Actual de la Construcción': 'Licitación',
                                            'Empresa a cargo de Corrientes Débiles': 'Por definir',
                                            'Estado de Avance - Corrientes Débiles': 'Pendiente',
                                            'Oportunidad / Acción a Tomar': f'Revisar si incluye llamado de enfermería. ID: {codigo}',
                                            'Contactos / Observaciones': f'https://www.mercadopublico.cl/DirectorioLicitacion/fichaLicitacion.aspx?codlic={codigo}'
                                        })
                                        ids_procesados.add(codigo)
                            break 
                        else:
                            print(f"Aviso API: {data.get('Mensaje', 'Ticket inválido')}")
                    else:
                        print(f"Error servidor MP (Status {response.status_code})")
                
                except Exception as e:
                    print(f"Error de conexión: {e}")
                    time.sleep(10)

        # --- 3. GENERACIÓN DEL EXCEL ---
        if proyectos_encontrados:
            df = pd.DataFrame(proyectos_encontrados)
            nombre_archivo = 'Seguimiento_Proyectos_Hospitalarios_Auto.xlsx'
            df.to_excel(nombre_archivo, index=False)
            print(f"¡Listo compita! Se encontraron {len(proyectos_encontrados)} licitaciones de salud y se guardaron en {nombre_archivo}.")
        else:
            print("No se encontraron licitaciones nuevas de hospitales en estos últimos 3 días.")

    except Exception as e:
        print(f"Error crítico: {e}")

if __name__ == "__main__":
    buscar_hospitales_excel()
