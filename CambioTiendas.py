import pandas as pd
import psycopg2
import datetime

# ---VARIABLES---#
Tiendas = ["BK", "LC", "CH", "PC", "PP", "CC", "CK", "DD", "BK-CALL",
           "LC-CALL", "CH-CALL", "PC-CALL", "PP-CALL", "CC-CALL", "CK-CALL", "DD-CALL",]

# ---FUNCIONES BASE---#

# ---Convierte el Reporte de Excel en un dataframe---#
def leer_Reporte(nombreReporte):
    try:
        df_reporte = pd.read_excel(nombreReporte, engine='xlrd')
        return df_reporte
    except:
        return "--- Archivo no encontrado"

# ---Convierte el archivo de FAC en un dataframe---#
def leer_FAC(nombreFac):
    try:
        df_reporte = pd.read_excel(nombreFac, engine='xlrd')
        return df_reporte
    except:
        return "--- Archivo no encontrado"

# ---obtiene los numeros de autorizacion---#
def obtener_authNum(numFila, dfReporte):

    authNum = dfReporte.at[numFila, 'OBS']

    # ---Verificar si la celda no esta vacia---#
    if pd.isna(authNum) == True:
        pass
    else:

        # ---Verifica si tiene numero de autorizacion---#
        if authNum[-6:] != "CION: ":
            return authNum[-6:]
        else:
            pass

# ---Obtener el OrderId segun el authNum---#
def Obtener_orderId(authNum, dfFAC):

    try:
        index = dfFAC[dfFAC["Auth Code"] == authNum].index[0]
        OrderId = dfFAC.at[index, "Order ID"]

        return int(OrderId)
    except:
        # print("!!!Auth_Num "+str(auth_Num)+ " no encontrado en FAC")
        return "!!!Auth_Num "+str(auth_Num) + " no encontrado en FAC"

# ---Concectar a la DB y obtener tienda---#
def Obtener_Tienda(orderId):

    if orderId >= 1000000:
        orderFull = "000" + str(orderId)
    else:
        orderFull = "0000" + str(orderId)

    query = "SELECT cmbpa.username as Tienda FROM public.csm_prc_order cpo LEFT JOIN cs_mst_business_partner_address cmbpa ON cmbpa.id = CAST(cpo.provider AS INTEGER) WHERE cpo.code = '" + str(
        orderFull) + "'"

    # ---obtener tienda segun orderId---#
    try:

        if orderId is None:
            pass
        else:

            cursor.execute(str(query))
            tienda = cursor.fetchone()[0]

            if tienda[:2] == "IJ":
                return "CL" + str(tienda[-2:])
            else:
                return str(tienda)

    except:
        return "No se pudo obtener tienda de la orden: " + str(orderId)

# ---------------------------------------------------------------------------#
# ---------------------------------------------------------------------------#


# ---Obtener nombres de los archivos---#
print("Code by Rofer")
Reporte = input("Ingrese el nombre del reporte: ") + ".xls"
FAC = input("Ingrese el nombre del archivo de FAC:  ")+".xls"

# ---Crear el .txt para el log---#
fecha_actual = datetime.datetime.today().strftime('%Y-%m-%d')
Log = f"Tiendas_{fecha_actual}.txt"

# ---DB data---#
host = 'avecsm.cksl0gxfpv9j.us-east-2.rds.amazonaws.com'
database = 'AVECSM'
user = 'itintur'
password = 'Intur#2021!'


# ---Conectar a la DB---#
conecction = psycopg2.connect(
    host=host, database=database, user=user, password=password)
cursor = conecction.cursor()


df_Reporte = leer_Reporte(Reporte)
df_FAC = leer_FAC(FAC)

i = 0

if isinstance(df_Reporte, str) or isinstance(df_FAC, str):
    print("---Archivo no encontrado")
    input("Presiona Enter para salir...")

else:
    with open(Log, "w") as archivo:
        pass

    while i < df_Reporte.shape[0]:

        auth_Num = obtener_authNum(i, df_Reporte)

        if auth_Num is None:
            pass
        else:

            Order_id = Obtener_orderId(auth_Num, df_FAC)
            if isinstance(Order_id, str):
                print(Order_id)
                with open(Log, "a") as archivo:
                    archivo.write(Order_id + "\n")

                pass
            else:

                Tienda = Obtener_Tienda(Order_id)
                print(auth_Num + " / " + str(Order_id) + " / " + Tienda)

                with open(Log, "a") as archivo:
                    archivo.write(auth_Num + " / " +
                                  str(Order_id) + " / " + Tienda + "\n")

                # ---Cambiar Tienda---#
                df_Reporte.at[i, "RESTAURANTE"] = Tienda
                df_Reporte.at[i, "OBS"] = ""

        i += 1
    df_Reporte.to_excel(Reporte[:-4]+"Nuevo"+".xlsx", index=False)
    cursor.close()
    conecction.close()
    input("Presiona Enter para salir...")
