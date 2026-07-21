import os
import duckdb
import msal
import requests

# ====================================================================
# 1. RECUPERAR CREDENCIALES SEGURAS DESDE GITHUB SECRETS
# ====================================================================
# En lugar de escribir las contraseñas aquí, Python las toma de la 
# caja fuerte digital de GitHub que configuramos en el Paso 2.
TENANT_ID = os.environ.get('TENANT_ID')
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')

# Definición de rutas y nombres de archivos
URL_MEF = "https://fs.datosabiertos.mef.gob.pe/datastorefiles/2026-Gasto-Diario.csv"
OUTPUT_CSV = "mef_gasto_maestro_filtrado.csv"
DB_NAME = "mef_analisis_final.duckdb"
ARCHIVO_RESUMEN = "resumen_ejecucion_gobierno_nacional.csv"

# Sectores que deseas analizar para el MINAM y los demás ministerios
SECTORES_OBJETIVO = (
    "VIVIENDA CONSTRUCCION Y SANEAMIENTO", "DEFENSA", "RELACIONES EXTERIORES",
    "SALUD", "PRODUCCION", "PRESIDENCIA CONSEJO MINISTROS",
    "ENERGIA Y MINAS", "JUSTICIA", "ECONOMIA Y FINANZAS", "INTERIOR",
    "MUJER Y POBLACIONES VULNERABLES", "TRABAJO Y PROMOCION DEL EMPLEO",
    "TRANSPORTES Y COMUNICACIONES", "DESARROLLO E INCLUSION SOCIAL",
    "EDUCACION", "AMBIENTAL", "AGRARIO Y DE RIEGO", "CULTURA",
    "COMERCIO EXTERIOR Y TURISMO"
)

# ====================================================================
# 2. DESCARGA ACELERADA DEL MEF EN LA NUBE DE GITHUB
# ====================================================================
print("🚀 [1/3] Descargando archivo masivo del MEF en la nube...")
# Usamos aria2c para exprimir al máximo la velocidad de los servidores de GitHub
os.system(f"aria2c -x 16 -s 16 -k 1M --file-allocation=none --allow-overwrite=true -d . -o {OUTPUT_CSV} {URL_MEF}")

# ====================================================================
# 3. PROCESAMIENTO INTELIGENTE CON DUCKDB
# ====================================================================
print("\n⚙️ [2/3] Procesando, filtrando y agrupando datos con DuckDB...")
if os.path.exists(DB_NAME):
    os.remove(DB_NAME)

con = duckdb.connect(DB_NAME)
con.execute("PRAGMA threads=4;") # Usa múltiples núcleos de la nube

# Filtramos al vuelo solo el Gobierno Nacional ('E') para no saturar memoria
con.execute(f"""
    CREATE TABLE gasto_nacional AS
    SELECT * 
    FROM read_csv_auto('{OUTPUT_CSV}', nullstr=' ', ignore_errors=true, strict_mode=false)
    WHERE TRIM(NIVEL_GOBIERNO) = 'E';
""")

# Consulta de agregación (PIM, Devengado y % de Ejecución)
query = """
    SELECT
        SECTOR AS COD_SECTOR,
        SECTOR_NOMBRE,
        TIPO_ACT_PROY_NOMBRE,
        SUM(TRY_CAST(MONTO_PIM AS DOUBLE)) AS PIM_TOTAL,
        SUM(TRY_CAST(MONTO_DEVENGADO AS DOUBLE)) AS DEVENGADO_TOTAL,
        ROUND((SUM(TRY_CAST(MONTO_DEVENGADO AS DOUBLE)) / NULLIF(SUM(TRY_CAST(MONTO_PIM AS DOUBLE)), 0)) * 100, 2) AS PORCENTAJE_EJECUCION
    FROM gasto_nacional
    WHERE UPPER(TIPO_ACT_PROY_NOMBRE) IN ('ACTIVIDAD', 'PROYECTO')
      AND UPPER(TRIM(SECTOR_NOMBRE)) IN $sectores
    GROUP BY SECTOR, SECTOR_NOMBRE, TIPO_ACT_PROY_NOMBRE
    ORDER BY TIPO_ACT_PROY_NOMBRE, SECTOR;
"""

df_resumen = con.execute(query, {"sectores": SECTORES_OBJETIVO}).df()
con.close()

# Formateo de montos y guardado del CSV liviano de resumen
df_resumen['PIM_TOTAL'] = df_resumen['PIM_TOTAL'].map('{:,.2f}'.format)
df_resumen['DEVENGADO_TOTAL'] = df_resumen['DEVENGADO_TOTAL'].map('{:,.2f}'.format)
df_resumen.to_csv(ARCHIVO_RESUMEN, index=False, encoding='utf-8-sig')
print(f"✅ Archivo resumen generado con éxito.")

# ====================================================================
# 4. CONEXIÓN Y SUBIDA AUTOMÁTICA A SHAREPOINT
# ====================================================================
print("\n☁️ [3/3] Conectando con SharePoint para actualizar Power BI...")
autoridad = f'https://login.microsoftonline.com/{TENANT_ID}'
app = msal.ConfidentialClientApplication(CLIENT_ID, authority=autoridad, client_credential=CLIENT_SECRET)
respuesta_token = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])

if 'access_token' in respuesta_token:
    access_token = respuesta_token['access_token']
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Localizamos tu sitio de SharePoint institucional
    site_endpoint = "https://graph.microsoft.com/v1.0/sites/datascientist23.sharepoint.com:/sites/BASEDEDATOS-DASHBOARDMINAM"
    res_site = requests.get(site_endpoint, headers=headers)

    if res_site.status_code == 200:
        site_id = res_site.json().get('id')
        endpoint_subida = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ARCHIVO_RESUMEN}:/content"
        headers_subida = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/octet-stream'}

        with open(ARCHIVO_RESUMEN, 'rb') as archivo:
            datos_archivo = archivo.read()

        res_subida = requests.put(endpoint_subida, headers=headers_subida, data=datos_archivo)

        if res_subida.status_code in [200, 201]:
            print("🎉 ¡ÉXITO TOTAL! El archivo se subió a SharePoint y Power BI ya puede actualizarse.")
        else:
            raise Exception(f"❌ Error al subir a SharePoint: {res_subida.status_code} - {res_subida.json()}")
    else:
        raise Exception(f"❌ No se pudo encontrar el sitio de SharePoint: {res_site.status_code}")
else:
    raise Exception("❌ Error de autenticación con Azure. Revisa tus credenciales en los Secrets.")
