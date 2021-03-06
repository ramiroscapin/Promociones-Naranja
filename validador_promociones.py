from tkinter import *
from tkinter import font
from tkinter import filedialog
from tkinter.filedialog import askopenfile 

#Cramos raiz o root de nuestra aplicacion
ventana = Tk()
ventana.title("Validador Promociones")
ventana.geometry("740x470")
ventana.resizable(False,False)
ventana.configure(background="#FF7207")


color_boton = "#FF9E1B"
operador = ""
texto_pantalla = StringVar()

def clear():
    global operador
    operador = ""
    texto_pantalla.set(operador)
    
def mfileopen():
    file1 = filedialog.askopenfile()
    ruta_excel = ((str(file1)).split("'"))[1]
    label = Label(text=file1)
    pantalla_ruta_excel.insert(0,ruta_excel)
    return(((str(file1)).split("'"))[1])
        

def correr_programa():
    from datetime import datetime, date, time, timedelta
    import calendar
    from pprint import pprint
    import pandas as pd
    def eliminar_duplicados(lista_con_duplicados):
        lista_sin_duplicados = list(dict.fromkeys(lista_con_duplicados))
        return lista_sin_duplicados

    def levantar_excel():
        
        try:
            #nombre_excel = str(input("Escriba el Nombre del Excel: ")) + ".xlsx"
            #nombre_hoja = str(input("Escriba el Nombre de la Hoja: "))
            ###################### BORRAR
            nombre_excel =pantalla_ruta_excel.get()
            nombre_hoja = str(txt_hoja_excel.get())
            ###################### BORRAR
            xls = pd.read_excel(nombre_excel,sheet_name=nombre_hoja)
            xls_desplegable = pd.read_excel(nombre_excel,sheet_name='Desplegable')
            return xls, xls_desplegable
        except Exception as e:
            print("Nombre de excel o de hoja ERRONEOS: ",e)
            print("Si los datos son correctos, asegurese de que el excel esté en la misma carpeta que el archivo validador_promociones.py")
            return pd.DataFrame(), pd.DataFrame()

    def revisar_istitle(excel, errores_list, nombre_columna):
        # Revisa si los items de la columna son tipo "Titulo De Noticia"
        dic = {}
        columna = excel[nombre_columna]
        for index, item in enumerate(columna, start = 1):
            item = str(item).strip()
            if not item.istitle(): #and item!=nan:
                if item not in dic.keys():
                    dic[item]=[index]
                else:
                    dic[item].append(index)

        errores_list.append({nombre_columna:dic}) 
        return errores_list

    def revisar_desplegable(xls, desplegable, errores_list, columna_desplegable, columna_xls):
        dic = {}
        desplegable_columna = desplegable[columna_desplegable].dropna()
        # Quito los espacios al comienzo y al final
        trim_strings = lambda x: x.strip() if isinstance(x, str) else x 
        desplegable_columna = desplegable_columna.map(trim_strings)

        set_columna_desplegable = set(desplegable_columna)
        columna = xls[columna_xls]

        for index, item in enumerate(columna, start = 1):
                item = str(item).strip()
                if item not in set_columna_desplegable:# and item!=nan:
                    if item not in dic.keys():
                        dic[item]=[index]
                    else:
                        dic[item].append(index)
    
        if columna_xls == 'APLICACIÓN DEL DESCUENTO':
            try:
                del dic["nan"]
            except:
                pass

        errores_list.append({columna_xls:dic})
        return errores_list

    def concatenar_columnas(base, nombre_columna, nombre_columna1):
        nombre_columna =base[nombre_columna]
        nombre_columna1 =base[nombre_columna1]
    
    
        base["columna_concatenada"] = nombre_columna + " " + nombre_columna1    
        return(base["columna_concatenada"])

    def revisar_nombre_fantasia(xls, errores_list):
        # NOMBRE DE FANTASÍA: 
        # Deben estar escrito con la primera letra en mayúscula y el resto en minúscula, por cada palabra. 
        # Ejemplo: Casa Del Audio 
        nombre_columna = 'NOMBRE DE FANTASÍA' 
        errores_list = revisar_istitle(xls, errores_list, nombre_columna)
        ### TEST 
        #df = xls[not xls['NOMBRE DE FANTASÍA'].istitle()]['NOMBRE DE FANTASÍA']
        return errores_list

    def revisar_rubro(xls, desplegable, errores_list):
        # RUBRO: Debe coincidir con la LISTA DESPLEGABLE. 
        # Se puede usar la validación de datos o control de duplicados 
        # No deben quedar celdas sin información.
        nombre_col_desplegable = 'RUBROS ( DATA WAREHOUSE)'
        nombre_col_xls = 'RUBRO'
        errores_list = revisar_desplegable(xls, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def revisar_provincia(excel, desplegable, errores_list):
        # PROVINCIA: Debe coincidir con la LISTA DESPLEGABLE. 
        # Se puede usar la validación de datos o control de duplicados. 
        # No deben quedar celdas sin información.
        nombre_col_desplegable = 'PROVINCIA (Sin duplicados)'
        nombre_col_xls = 'PROVINCIA'
        errores_list = revisar_desplegable(excel, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def revisar_localidad(excel, desplegable, errores_list):
        nombre_col_desplegable = 'LOCALIDADES (Sin duplicados)'
        nombre_col_xls = 'LOCALIDAD'
        errores_list = revisar_desplegable(excel, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def revisar_direccion(excel, errores_list):
        # DIRECCIÓN: Deben estar escrito con la primera letra en mayúscula y el resto en minúscula, 
        # por cada palabra. 
        # Ejemplo: Av. Corrientes 3135  
        nombre_columna = 'DIRECCIÓN '
        errores_list = revisar_istitle(excel, errores_list, nombre_columna)
        return errores_list

    def revisar_plan_ppal(excel, desplegable, errores_list):
        # PLAN PRINCIPAL: Debe coincidir con la LISTA DESPLEGABLE. 
        # Se puede usar la validación de datos o control de duplicados. 
        # No deben quedar celdas sin información.  
        nombre_col_desplegable = 'PLAN PRINCIPAL'
        nombre_col_xls = 'PLAN PRINCIPAL'
        errores_list = revisar_desplegable(excel, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def revisar_cero_interes(df, errores_list):
        """
        Devuelve los errores de las columnas CFT / TEA / TNA
        SI PLAN ES CERO INTERES DEBEN SER 0 
        """
        dic = {}
        # Tomo todas las que tengan cero interés
        cero_int = df[df['PLAN PRINCIPAL'].str.contains("cero int", na=False)]
        # Si son de cero interés, y las columnas no son 0, es un error.
        errores_CFT = cero_int.loc[(cero_int['CFT'] != 0)]
        errores_TEA = cero_int.loc[(cero_int['TEA'] != 0)]
        errores_TNA = cero_int.loc[(cero_int['TNA'] != 0)]
        # Tomo los índices y elimino duplicados
        lista_indices = (list((errores_CFT.index)+1)) + (list((errores_TEA.index)+1)) + (list((errores_TNA.index)+1))
        lista_indices = eliminar_duplicados(lista_indices)
        dic["Plan sin interés pero NO SON 0.0%"] = lista_indices
        errores_list.append({"CFT / TEA / TNA":dic})
        return errores_list

    def revisar_cuotas_fijas(df, errores_list):
        """
        Devuelve los errores de las columnas CFT / TEA / TNA
        SI PLAN ES DE CUOTAS FIJAS NO DEBEN SER 0 % 
        """
        dic = {}
        # Tomo todas las que tengan cero interés
        cuotas_fijas = df[df['PLAN PRINCIPAL'].str.contains("cuotas fijas", na=False)]
        # Si son de cero interés, y las columnas no son 0, es un error.
        errores_CFT = cuotas_fijas.loc[(cuotas_fijas['CFT'] == 0)]
        errores_TEA = cuotas_fijas.loc[(cuotas_fijas['TEA'] == 0)]
        errores_TNA = cuotas_fijas.loc[(cuotas_fijas['TNA'] == 0)]
        # Tomo los índices y elimino duplicados
        lista_indices = (list((errores_CFT.index)+1)) + (list((errores_TEA.index)+1)) + (list((errores_TNA.index+1)))
        lista_indices = eliminar_duplicados(lista_indices)
        dic["Plan cuotas fijas pero SON 0.0%"] = lista_indices

        errores_list.append({"CFT / TEA / TNA":dic})
        return errores_list

    def revisar_interes(excel, desplegable, errores_list):
        # CFT / TEA / TNA: 
        # No deben quedar celdas sin información. 
        # Si se trata de un plan cero interés el costo debe ser 0,00%, 
        # Si se trata de plan con cuota fija se debe informar el costo.
    
        columnas = ['PLAN PRINCIPAL', 'CFT', 'TEA', 'TNA']
        selec = excel[columnas]
        errores_list = revisar_cero_interes(selec, errores_list)
        errores_list = revisar_cuotas_fijas(selec, errores_list)
    
        return errores_list
    
    def revisar_descuento(excel, desplegable, errores_list):
        # DESCUENTO U OBSEQUIO PRINCIPAL: 
        # Debe coincidir con la LISTA DESPLEGABLE. 
        # No deben quedar celdas sin información
        nombre_col_desplegable = 'DESCUENTO / OBSEQUIO PRICIPAL'
        nombre_col_xls = 'DESCUENTO U OBSEQUIO PRINCIPAL'
        errores_list = revisar_desplegable(excel, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def aplicacion_descuento(excel, desplegable, errores_list):
        # DESCUENTO U OBSEQUIO PRINCIPAL: 
        # Debe coincidir con la LISTA DESPLEGABLE. 
        # No deben quedar celdas sin información
        nombre_col_desplegable = 'APLICACIÓN DESCUENTO'
        nombre_col_xls = 'APLICACIÓN DEL DESCUENTO'
        errores_list = revisar_desplegable(excel, desplegable, errores_list, nombre_col_desplegable, nombre_col_xls)
        return errores_list

    def verificar_sin_descuento_y_nulo_en_aplicacion_descuento(excel, errores_list):
        dic = {}
        columna_con_nulos_ok = excel["APLICACIÓN DEL DESCUENTO"]
        columna_sin_descuento = excel["DESCUENTO U OBSEQUIO PRINCIPAL"]
        indices_nulos = []
        indices_sin_descuento=[]
        dic["Nulos en filas de la columna APLICACIÓN DEL DESCUENTO y no tengo escrito sin descuento en columna DESCUENTO U OBSEQUIO PRINCIPAL "] = []
    
        for index, item in enumerate(columna_con_nulos_ok.isnull(), start = 1):
            if item == True:
                indices_nulos.append(index)
        for index, item in enumerate(columna_sin_descuento, start = 1):
            if item == "Sin descuento":
                indices_sin_descuento.append(index)
                if index not in indices_nulos:
                    if item not in dic.keys():
                        dic[item] = [index]
                    else:
                        dic[item].append(index)
        for item in indices_nulos:
            if item not in indices_sin_descuento:
                dic["Nulos en filas de la columna APLICACIÓN DEL DESCUENTO y no tengo escrito sin descuento en columna DESCUENTO U OBSEQUIO PRINCIPAL "].append(item)
        errores_list.append({"verificar combinacion entre columna (DESCUENTO U OBSEQUIO PRINCIPAL)  y columna (APLICACIÓN DEL DESCUENTO)  ":dic})
        return errores_list

    def revisar_nroCA(excel, errores_list):
        # Verificar que los números de CA poseen 9 dígitos y que no tengan "." ni "/". 
        # No deben quedar celdas vacías, si esto sucede 
        # la promoción no se va a incluir en el motor de recomendación
        dic = {}
        columna = excel['NRO. DEL CA']

        indexes_num, indexes_notnum = [], [] 
        for index, item in enumerate(columna, start = 1):
            try:
                item = int(item) # si no es int, sale por except
                item = str(item) # para poder contar la longitud
                if len(item) != 9:
                    # ERROR!
                    # Es un entero pero no tiene longitud de 9 dígitos
                    indexes_num.append(index)
            except: # ERROR!
                # No se pudo transformar a INT
                # No es dígito
                indexes_notnum.append(index)
        dic = {"es número pero no tiene 9 dígitos": indexes_num,"no es un número número de 9 dígitos" : indexes_notnum}

        errores_list.append({'NRO. DEL CA':dic}) 
        return errores_list

    def verificar_provincias_localidades(excel,desplegable,errores_list):
        dic = {}
        #desplegable
        columna_a_comparar = (concatenar_columnas(desplegable,'PROVINCIAS','LOCALIDADES')).dropna()
        trim_strings = lambda x: x.strip() if isinstance(x, str) else x
        columna_a_comparar = columna_a_comparar.map(trim_strings)
    
        set_columna_a_comparar = set(columna_a_comparar)
        #excel
        columna = concatenar_columnas(excel,'PROVINCIA','LOCALIDAD')
    
        for index, item in enumerate(columna, start = 1):
            item = str(item).strip()
            if item not in set_columna_a_comparar:
                if item not in dic.keys():
                    dic[item]=[index]
                else:
                    dic[item].append(index)
        errores_list.append({"Combinación Provincias y Localidades":dic})
        return(errores_list)

    def revisar_vigencia(xls, errores_lista):
        dic = {}
        columna = excel['VIGENCIA DESDE']
    #     mes_actual = 12
        mes_actual = (datetime.now()).month
        un_mes_mas_del_actual = mes_actual + 1
        dos_meses_mas_del_actual =mes_actual + 2
    #     año_actual = 2020
        año_actual = (datetime.now()).year
        un_año_mas_del_actual = año_actual +1
    
        for index,fecha in enumerate(columna,start = 1):
            ## REVISO EL TYPE DE LA FECHA, YA QUE SI HAY ALGUNA ENTRADA QUE NO SEA FECHA ME MUESTRE COMO ERROR
            # A REVISAR 
            if type(fecha) == int or type(fecha) == float or type(fecha) == str:
                if fecha not in dic.keys():
                    dic[fecha] = [index]
                else:
                    dic[fecha].append(index)
            else:
                #### LA FECHA ANALIZADA TIENE QUE SER MAYOR O IGUAL AL MES ACTUAL Y MENOR A DOS MESES MAS ADELANTE DEL ACTUAL
                if fecha.month >= mes_actual and fecha.month < dos_meses_mas_del_actual and año_actual == fecha.year:
                    pass
                else:
                    if fecha.month == 1 and fecha.year == un_año_mas_del_actual:
                        pass
                    else:
                        if fecha not in dic.keys():
                            dic[fecha] = [index]
                        else:
                            dic[fecha].append(index)
                        
        errores_list.append({"columna_desde":dic})
        return errores_list

    def revisar_columa_tope_reintegro(excel, errores_list):
        dic = {}
        columna_tope_reintegro = excel["TOPE DE REINTEGRO"]
        for index, item in enumerate(columna_tope_reintegro):
            if type(item) == int:
                pass
            else:
                if item not in dic.keys():
                    dic[item] = [index]
                else:
                    dic[item].append(index)
        errores_list.append({"tope reintegro":dic})
        return errores_list

    def escribir_errores(errores_list):
        with open("borrar.txt", "w") as output:
            for item_columna in errores_list:
                nombre_columna = list(item_columna.keys())[0]
                #print(nombre_columna)
                a = f"Para la columna {nombre_columna} los errores son:\n"
                output.write(a)
                output.write('\n')
                for key, value in item_columna[nombre_columna].items():
                    if key == 'nan':
                        key = 'VACIO'
                    b = f"El error '{key}' se repite en las filas: {value}\n"
                    output.write(b)
                    output.write('\n')
                            
                output.write('------------------------------------------------------------------------------\n')

            #output.write(str(errores_list))
            print("Escrito ERRORES en borrar.txt")
            operador = "Errores escritos en archivo borrar.txt"
            texto_pantalla.set(operador)
            return 

    
    
    
    if __name__ == "__main__":
        excel, desplegable = levantar_excel()
        if excel.empty:
            print("Excel Vacío: Finalizando NOK")
        else:
            errores_list = revisar_nombre_fantasia(excel, [])
            errores_list = revisar_rubro(excel, desplegable, errores_list)
            errores_list = revisar_provincia(excel, desplegable, errores_list)
            errores_list = revisar_localidad(excel, desplegable, errores_list)
            errores_list = verificar_provincias_localidades(excel,desplegable,errores_list)
            ############ REVISAR PROVINCIA Y LOCALIDAD
            #errores_list = revision_prov_localidad(excel, desplegable, errores_list)
        
            errores_list = revisar_direccion(excel, errores_list)
            errores_list = revisar_plan_ppal(excel, desplegable, errores_list)
            errores_list = revisar_interes(excel, desplegable, errores_list)
            errores_list = revisar_descuento(excel, desplegable, errores_list)
            errores_list = verificar_sin_descuento_y_nulo_en_aplicacion_descuento(excel, errores_list)
            # no se hace una revisión cruzada con el descuento
            errores_list = aplicacion_descuento(excel, desplegable, errores_list)
            errores_list = revisar_nroCA(excel, errores_list)
            errores_list = revisar_vigencia(excel, errores_list)
        
            #REVISAR TOPE REINTEGRO
            ## ME SIGUE TOMANDO NULOS Y EN EXCEL CUANDO APARECEN NUMEROS CON DECIMALES TERMINADOS EN 0 EJ: 1.0 , 2.0 , 3.0
            ## ME LOS CARGA SIN EL .0 LO CONSIDERA COMO INT.
            errores_list = revisar_columa_tope_reintegro(excel, errores_list)


            escribir_errores(errores_list)
        
            """
            revisar_legales()
            revisar_dias()
            revisar_vigencia()
            revisar_tope_reintegro()
            revisar_especificacion() """
    

imagen = PhotoImage(file="fondo3.png")    
background = Label(image = imagen)
background.place(x = 0, y = 0, relwidth = 1, relheight = 1)


txt_hoja_excel = Entry(ventana,width = 57 , borderwidth = 8, background = '#F0EBE6')
txt_hoja_excel.place(x=355,y=260)


boton_cargar_excel = Button(ventana, text = "Buscar Excel", bg=color_boton , width = 50, height = 2, command = mfileopen)
boton_cargar_excel.place(x=355 , y=110)
pantalla_ruta_excel = Entry(ventana,width = 57, borderwidth = 8, background = '#FFCE8F')
pantalla_ruta_excel.place(x=355,y=165)

boton_cargar = Button(ventana, text = "Correr Programa !", bg=color_boton , width = 50, height = 2, command = correr_programa)
boton_cargar.place(x=355 , y=330)

# Ejecutamos el bucle infinito
ventana.mainloop()
