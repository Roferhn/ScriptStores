import pandas as pd
import psycopg2




# ---VARIABLES---#
Tiendas = ["BK", "LC", "CH", "PC", "PP", "CC", "CK", "DD", "BK-CALL", "LC-CALL", "CH-CALL", "PC-CALL", "PP-CALL", "CC-CALL", "CK-CALL", "DD-CALL",]

# ---FUNCIONES BASE---#

# ---Convierte el Reporte de Excel en un dataframe---#
def leer_Reporte(nombreReporte):
    try:
        df_reporte = pd.read_excel(nombreReporte, engine='xlrd')
        return df_reporte
    except:
        return "---","Archivo no encontrado"

# ---Convierte el archivo de FAC en un dataframe---#
def leer_FAC(nombreFac):
    try:
        df_reporte = pd.read_excel(nombreFac)
        return df_reporte
    except:
        return "---","Archivo no encontrado"

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

#---Obtener el OrderId segun el authNum---#
def Obtener_orderId(authNum, dfFAC):

    try:
        index= dfFAC[dfFAC["Auth Code"] == authNum].index[0]
        OrderId= dfFAC.at[index, "Order ID"]
        return OrderId
    except:
        print("!!!","Auth_Num ", auth_Num, " no encontrado en FAC")
        return False

#---Concectar a la DB y obtener tienda---#
def Obtener_Tienda(orderId):
    host = 'avecsm.cksl0gxfpv9j.us-east-2.rds.amazonaws.com'
    database =  'AVECSM'
    user = 'itintur'
    password = 'Intur#2021!'

    #---Conectar a la DB---#
    conecction = psycopg2.connect(host=host, database=database, user=user, password=password)
    cursor = conecction.cursor()

    if orderId >= 1000000:
        orderFull = "000" + str(orderId)
        print(orderFull)
    else: 
        orderFull = "0000" + str(orderId)
        print(orderFull)

    query= "SELECT cmbpa.username as Tienda FROM public.csm_prc_order cpo LEFT JOIN cs_mst_business_partner_address cmbpa ON cmbpa.id = CAST(cpo.provider AS INTEGER) WHERE cpo.code = '" + str(orderFull) + "'"

    #---obtener tienda segun orderId---#
    try:

        if orderId is None:
            pass
        else:

            cursor.execute(str(query))
            tienda = cursor.fetchone()[0]

            cursor.close()
            conecction.close()

            if tienda[:2] == "IJ":
                return "CL" + str(tienda[-2:])
            else:
                return str(tienda)
    
    except:
        return "No se pudo obtener tienda de la orden: " + str(orderId)



# ---------------------------------------------------------------------------#
print("Code by Rofer")
#Reporte = "VNP-05072023" + ".xls"
Reporte = input("Nombre del reporte?") + ".xls"

#FAC = "Transactions_20230705_135000"+".xlsx"
FAC = input("Nombre del archivo de FAC?")+".xls"
i = 0
df_Reporte = leer_Reporte(Reporte)
df_FAC = leer_FAC(FAC)

while i < df_Reporte.shape[0]:

    auth_Num = obtener_authNum(i, df_Reporte)

    if auth_Num is None:
        pass
    else:
        Order_id = Obtener_orderId(auth_Num,df_FAC)


        if Order_id==False:
            pass
        else:
            Tienda = Obtener_Tienda(Order_id)
            print(auth_Num + " / " + str(Order_id) + " / " + Tienda)

            #---Cambiar Tienda---#
            df_Reporte.at[i,"RESTAURANTE"] = Tienda
            df_Reporte.at[i,"OBS"] = ""

    i += 1
df_Reporte.to_excel(Reporte[:-4]+"Nuevo"+".xlsx",index=False)