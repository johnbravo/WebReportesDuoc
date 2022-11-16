import cx_Oracle #en cmd ejecutar pip install cx_oracle el cual nos permite instalar cx_oracle, este modulo nos permite ingresar a la base de datos oracle
#import numpy as np 
from io import BytesIO #en cmd ejecutar pip install requires.io , El módulo Python io nos permite administrar las operaciones de entrada y salida relacionadas con archivos, nos permitira exportar nuestro reporte en excel
import pandas as pd #en cmd ejecutar pip install pandas , libreria que nos permitira manipular y tratar datos , permitiendo crear los dataframes para imprimir excel



try:
    conexion=cx_Oracle.connect(  #cadena de conexion a la base de datos oracle.
    user='system',          #indicamos el usuario de la base datos
    password='admin',       #indicamos la contraseña de la base de datos
    dsn='10.155.7.242/xe')  #nombre de host del equipo , al ser una maquina virtual queda con una ip, este nombre de host debe coincidir con tsnames y listener

except Exception as err:
    print('Error en la conexion a la base de datos', err)   
else:
    print('conectado a Oracle Database', conexion.version)   #try que nos avisara cuando ocurra algun problema con la conexion    


    

# pip install flask, flask es la herramienta que nos permitira trabajar python junto con el html
#importamos flask
#importamos render template el cual nos permite dirigir nuestras variables a los diversos templates html 
#importamos request el cual nos permitira traer valores desde nuestro html y captarlos en donde trabajamos nuestro codigo en "default.py"
#importamos send_file el cual permite genear mediante el navegador web nuestro archivo en este caso excel
from flask import Flask, render_template, request , send_file ,jsonify # importamos los elementos de flask que necesitamos, esto se hace solo una vez
app = Flask(__name__)

#flask nos permite trabajar junto con el html de esta siguiente forma definimos 3 lineas de codigo que son las mas importantes
#@app.route("/") nos permite definir la ruta que tendra la pagina , en este caso sera la raiz "index"
#def home(): nos permite nombrar a nuestro pequeño fragmento de codigo el cual trabajara con nuestra pagina web
#return render_template() nos guiara nustras variables a dicho template en este caso "index.html"
@app.route("/", methods =["GET", "POST"]) 
def home():

    cur_01=conexion.cursor() #creamos el cursor

#realizaremos un select que devolvera cada 1 de los edificios con su respectivo nombre
    select_datos= """ select * from edificio"""    #creamos la consulta y la insertamos en una variable
    cur_01.execute(select_datos) #mediante el cursor executamos la consulta
    row = cur_01.fetchall() #todos los datos conseguidos son insertados como tupla en la variable
    idedificio = [item[0] for item in row] #luego cada item de la tupla lo descomponemos y lo almacenamos en variables independientes
    Nedificio = [item[1] for item in row]

    


    conexion.commit()

    return render_template("index.html",  len = len(row),lenideficiio = len(idedificio),lenNedificio = len(Nedificio),
                                          row = row,idedificio = idedificio,Nedificio = Nedificio)

#@app.route('/gfg', methods =["GET", "POST"]) 
#en este caso nos guiara a la siguiente pagina y en cuanto a get o post , usaremos post para pedir al formulario las diversas variables
@app.route('/Formulario', methods =["GET", "POST"])
def Formulario():
#a continuacion recogeremos la variable seleccionada en el form anterior para filtrar todas las salas en base a que edificio se selecciono, y escoger 1 sala
#ademas seleccionaremos el tipo de problema 
    if request.method == "POST":
     nedificiohtml = request.form.get("nedificio")#obtenemos el valor seleccionado en la pagina anterior y lo usamos en la consulta siguiente

    
    cur_edi = conexion.cursor()
        #seleccionamos el nombre del edificio para mostrarlo
    select_edi = (""" select nombreedificio from edificio where idedificio= :var2""") #var2 es la variable obtenida previamente del form anterior
    cur_edi.execute(select_edi, {'var2': nedificiohtml}) #aca usamos la variable recogida 
    rowedi = cur_edi.fetchall()
    nombreedificio = [item[0] for item in rowedi]


    cur_d = conexion.cursor()
        #seleccionamos la sala para mostrarla
    select_d = (""" select idsala,numerosala from sala where fk_edificio = :var order by idsala asc""")
    cur_d.execute(select_d, {'var': nedificiohtml})
    row2 = cur_d.fetchall()
    idsala = [item[0] for item in row2]
    nsalahtml = [item[1] for item in row2]


    cur_tipoproblema = conexion.cursor()
        #seleccionamos el problema para mostrarlo
    select_datos_tipo_problema = (""" select idproblema,tipoproblema from tipoproblema""")
    cur_tipoproblema.execute(select_datos_tipo_problema)
    row4 = cur_tipoproblema.fetchall()
    idproblema = [item[0] for item in row4]
    tipoproblema = [item[1] for item in row4]

#si ocurre algun problema este fragmento nos permite edentificarlo
    
#enviamos todas las vairables al index2 mediante el render_template
    return render_template('index2.html',lensalahtml = len(nsalahtml),nsalahtml=nsalahtml,idsala=idsala,lentipoproblema=len(tipoproblema), tipoproblema=tipoproblema ,c=nedificiohtml,
                           lenidproblema=len(idproblema),idproblema=idproblema,lennombreedificio=len(nombreedificio),nombreedificio=nombreedificio)

#este fragmento de codigo peretenece a nuestro tecera pagina donde le haremos saber al usuario que un operador se dirigira a esta mencionando la sala en al cual
#indico el problema
@app.route('/final', methods =["GET", "POST"]) 
def final():
    if request.method == "POST":#obtenemos todas las variables del formulario anterior
        idedifi=request.form.get("idedifi")
        idsala = request.form.get("idsaladeclases")
        idproblema= request.form.get("Problema")
        mensaje=request.form.get("message")
#una vez las variables estan captadas las vamos ingresando en nuestra consulta la cual insertara en la base de datos nuestro reporte, creandoce el "reporte"
        try:
            cur_insert = conexion.cursor()

            insert_datos = (""" insert into reporte (descripcionproblema,fksala,fktipoproblema,fecha)values (:testing, :x, :y,sysdate)""")

            cur_insert.execute(insert_datos, {'testing': mensaje, 'x': idsala, 'y': idproblema})

            

        except Exception as err:
            print('error insertando datos', err)
        else:
            print('Datos insertandos correctamente!')
            conexion.commit()

#enviamos nuestras variables a final.html medietne render template 
    return render_template('final.html',idsala=idsala,idproblema=idproblema,idedifi=idedifi,mensaje=mensaje)

#------modulo de reportes--------
#el modulo de reportes es solo accesible al ingresar la ip del sevidor 10.20.4.56:88/reportefinal el cual nos muestra los ultimos 3 reportes
#ademas cada vez que se ingrese un reporte aparecera de color amarillo durante x cantidad de tiempo y junto con una alerta el nuevo reporte avisandole al operador

@app.route("/reportefinal")#nos genera la ruta /reportefinal
def reportefinal():#creamos el fragmento de codigo flask llamado reportefinal()
    try:
        cur_final = conexion.cursor()
        #seleccionamos el ultimo reporte lo mostramos por x cantidad de segundos y alertamos al usuario que alguien a generado 1 reporte
        select_datos_final = """ select id,descripcionproblema,fksala,tipoproblema,TO_CHAR (fecha, 'HH24:MI:SS'),nombreedificio from reporte 
                                    inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema
                                    inner join sala  on reporte.fksala = sala.numerosala
                                    inner join edificio on sala.fk_edificio = edificio.idedificio where fecha >= sysdate - 0.5/(24*60) order by id desc"""
        cur_final.execute(select_datos_final)
        rowfinal = cur_final.fetchall()
        idfinal = [item[0] for item in rowfinal]
        descripfinal = [item[1] for item in rowfinal]
        salafinal =[item[2] for item in rowfinal]
        problemafinal = [item[3] for item in rowfinal]
        fecha= [item[4] for item in rowfinal]
        edifn= [item[5] for item in rowfinal]

        cur_finalf = conexion.cursor()
        #seleccionamos los ultimos 3 reportes de la base de datos 
        select_datos_finalf = """ select id,descripcionproblema,fksala,tipoproblema,TO_CHAR (fecha, 'HH24:MI:SS'),nombreedificio from reporte 
                                    inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema
                                    inner join sala  on reporte.fksala = sala.numerosala
                                    inner join edificio on sala.fk_edificio = edificio.idedificio order by id desc FETCH FIRST 3 ROWS ONLY"""
        cur_finalf.execute(select_datos_finalf)
        rowfinalf = cur_finalf.fetchall()
        idfinalf = [item[0] for item in rowfinalf]
        descripfinalf = [item[1] for item in rowfinalf]
        salafinalf =[item[2] for item in rowfinalf]
        problemafinalf = [item[3] for item in rowfinalf]
        fechaf= [item[4] for item in rowfinalf]
        edificiof = [item[5] for item in rowfinalf]


    except Exception as err:
        print('error seleccionando datos', err)
    else:
        print('Datos seleccionados correctamente!')
        conexion.commit()
#enviamos mediante_render template los elementos 
    return render_template("reportefinal.html",lenidfinal=len(idfinal),idfinal=idfinal,descripfinal=descripfinal,salafinal=salafinal,problemafinal=problemafinal,fecha=fecha,
                                                lenidfinalf=len(idfinalf),idfinalf=idfinalf,descripfinalf=descripfinalf,salafinalf=salafinalf,problemafinalf=problemafinalf,fechaf=fechaf,edificiof=edificiof,edifn=edifn)
@app.route("/Reporteria")#generamos la ruta reporteria
def Reporteria():#creamos el nombre de nuestro modulo flask
#el modulo de reporteria cosiste en generar un reporte semanal de los problemas que se fueron generando durante esta
#las siguientes consultas consisten en filtrar mediante el tipo de problema y durante un plazo maximo los reportes 
    try:
        cur_rep = conexion.cursor()

        select_rep = (""" SELECT fksala ,COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema
                            where fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")
        cur_rep.execute(select_rep)
        rowrep = cur_rep.fetchall()
        salarep = [item[0] for item in rowrep]
        cantrep = [item[1] for item in rowrep]


        cur_proy = conexion.cursor()

        select_proy = (""" SELECT fksala ,COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Proyeccion' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_proy.execute(select_proy)
        rowproy = cur_proy.fetchall()
        salaproy = [item[0] for item in rowproy]
        cantproy = [item[1] for item in rowproy]



        cur_lentitud = conexion.cursor()

        select_lentitud = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Lentitud de equipo profesor' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_lentitud.execute(select_lentitud)
        rowlentitud = cur_lentitud.fetchall()
        salalentitud = [item[0] for item in rowlentitud]
        cantlentitud = [item[1] for item in rowlentitud]


        cur_lentitudlab = conexion.cursor()

        select_lentitudlab = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='problemas con equipos del laboratorio' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_lentitudlab.execute(select_lentitudlab)
        rowlentitudlab = cur_lentitudlab.fetchall()
        salalentitudlab = [item[0] for item in rowlentitudlab]
        cantlentitudlab = [item[1] for item in rowlentitudlab]


        cur_audio = conexion.cursor()

        select_audio = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Audio' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_audio.execute(select_audio)
        rowaudio= cur_audio.fetchall()
        salaaudio = [item[0] for item in rowaudio]
        cantaudio = [item[1] for item in rowaudio]


        cur_net = conexion.cursor()

        select_net = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='problemas de internet' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_net.execute(select_net)
        rownet= cur_net.fetchall()
        salanet = [item[0] for item in rownet]
        cantnet = [item[1] for item in rownet]

        cur_asis = conexion.cursor()

        select_asis = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Asistencia de un operador' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_asis.execute(select_asis)
        rowasis= cur_asis.fetchall()
        salaasis = [item[0] for item in rowasis]
        cantasis = [item[1] for item in rowasis]



        cur_carro = conexion.cursor()

        select_carro = (""" SELECT fksala, COUNT(fksala)
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Apertura de Carro' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_carro.execute(select_carro)
        rowcarro= cur_carro.fetchall()
        salacarro = [item[0] for item in rowcarro]
        cantcarro = [item[1] for item in rowcarro]

    except Exception as err:
        print('error seleccionando datos', err)
    else:
        print('Datos seleccionados correctamente!')
        conexion.commit()

#enviamos a nuestro html reporteria todos los valores obtenidos para generar nuestro reporte 
    return render_template("reporteria.html",lenrep=len(salarep),salarep=salarep,cantrep=cantrep,
                                            lenproy=len(salaproy),salaproy=salaproy,cantproy=cantproy,
                                            lensalalentitud=len(salalentitud),salalentitud=salalentitud,cantlentitud=cantlentitud,
                                            lensalalentitudlab=len(salalentitudlab),salalentitudlab=salalentitudlab,cantlentitudlab=cantlentitudlab,
                                            lensalaaudio=len(salaaudio),salaaudio=salaaudio,cantaudio=cantaudio,
                                            lensalanet=len(salanet),salanet=salanet,cantnet=cantnet,
                                            lensalaasis=len(salaasis),salaasis=salaasis,cantasis=cantasis,
                                            lensalacarro=len(salacarro),salacarro=salacarro,cantcarro=cantcarro)



@app.route('/return-files/')#generamos la ruta return-files el cual nos permitira enviar a este formulario los elementos a imprimir en excel
def return_files():#definimos nuestro fragmento de codigo flask con un nombre
    #seleccionamos la hora de la base de datos y se la ingresamos a una variable para usarla posteriormente en el nombre del archivo excel
    try:
        cur_fechaExcel = conexion.cursor()
        select_fechaExcel = (""" SELECT TO_CHAR (sysdate, 'DD.MM.YYYY') FROM DUAL """)

        cur_fechaExcel.execute(select_fechaExcel)
        rowfechaExcel= cur_fechaExcel.fetchall()
        fechaExcel = [item[0] for item in rowfechaExcel]
        
     #realizamos diversas consultas con los tipos de problemas para generar nuestro excel en base a los ultimos 7 dias.
     #cabe destacar que debe ser obligatorio que coincidan los nombres del select "sala" y "cantidad" a la hora de generar un dataframe o no
     #se imprimira como corresponde 
     #df_8 = pd.DataFrame(dataExcelcarro,columns = ['Sala','Cantidad']) ; las columnas del dataframe deben coincidir con las del select "Sala"y "Cantidad"
        cur_excel = conexion.cursor()
        select_excel = (""" SELECT fksala as Sala ,COUNT(fksala)as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema
                            where fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc """)

        cur_excel.execute(select_excel)
        data = cur_excel.fetchall()
       



        cur_excelproy = conexion.cursor()
        select_excelproy = (""" SELECT fksala as Sala,COUNT(fksala)as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Proyeccion' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc """)

        cur_excelproy.execute(select_excelproy)
        dataexcelproy= cur_excelproy.fetchall()

        cur_excelLenPro = conexion.cursor()
        select_excelLenPro = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Lentitud de equipo profesor' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc """)

        cur_excelLenPro.execute(select_excelLenPro)
        dataexcelLenPro= cur_excelLenPro.fetchall()


        cur_excelProLab = conexion.cursor()

        select_excelProLab = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='problemas con equipos del laboratorio' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_excelProLab.execute(select_excelProLab)
        dataexcelProLab = cur_excelProLab.fetchall()

        cur_excelAudio = conexion.cursor()

        select_excelAudio = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Audio' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_excelAudio.execute(select_excelAudio)
        dataexcelAudio = cur_excelAudio.fetchall()

        cur_Excelnet = conexion.cursor()

        select_Excelnet = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='problemas de internet' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_Excelnet.execute(select_Excelnet)
        dataExcelnet = cur_Excelnet.fetchall()

        cur_Excelasis = conexion.cursor()

        select_Excelasis = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Asistencia de un operador' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_Excelasis.execute(select_Excelasis)
        dataExcelasis = cur_Excelasis.fetchall()

        cur_Excelcarro = conexion.cursor()

        select_Excelcarro = (""" SELECT fksala as Sala,COUNT(fksala) as Cantidad
                            FROM reporte
                            inner join tipoproblema on reporte.fktipoproblema = tipoproblema.idproblema 
                            where tipoproblema='Apertura de Carro' and fecha > sysdate - 7 AND fecha <= sysdate 
                            GROUP BY fksala order by fksala asc""")

        cur_Excelcarro.execute(select_Excelcarro)
        dataExcelcarro= cur_Excelcarro.fetchall()
        
    
    except Exception as err:
        print('error consultando datos', err)
    else:
        print('Datos seleccionados correctamente!')
        conexion.commit()


    #creamos una variable y en esta almacenamos el dataframe,como primer parametro del dataframe recibe la tupla que obtuvo el cursor en la consulta
    #luego recibe las columnas que tendra el excel en este caso la sala y la cantidad de problemas reportados en esta.
    df_1 = pd.DataFrame(data,columns = ['Sala','Cantidad'])
    df_2 = pd.DataFrame(dataexcelproy,columns = ['Sala','Cantidad'])
    df_3 = pd.DataFrame(dataexcelLenPro,columns = ['Sala','Cantidad'])
    df_4 = pd.DataFrame(dataexcelProLab,columns = ['Sala','Cantidad'])
    df_5 = pd.DataFrame(dataexcelAudio,columns = ['Sala','Cantidad'])
    df_6 = pd.DataFrame(dataExcelnet,columns = ['Sala','Cantidad'])
    df_7 = pd.DataFrame(dataExcelasis,columns = ['Sala','Cantidad'])
    df_8 = pd.DataFrame(dataExcelcarro,columns = ['Sala','Cantidad'])
    #create an output stream
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    #taken from the original question
    df_1.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Resumen_General")
    df_2.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Proyeccion")
    df_3.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Lentitud_Equipo_Profesor")
    df_4.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Problema_Equipos_Laboratorio")
    df_5.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Audio")
    df_6.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Problemas_Internet")
    df_7.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Asistencia_Operador")
    df_8.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Apertura_Carros")
    workbook = writer.book
    worksheet = writer.sheets["Resumen_General"]
    format = workbook.add_format()
    format.set_bg_color('#eeeeee')
    worksheet.set_column(0,9,28)

    #the writer has done its job
    writer.close()

    #go back to the beginning of the stream
    output.seek(0)

    extension = '.xlsx'

    fecha = fechaExcel

    final = str(fecha)+extension 

    
    return send_file(output, download_name = final, as_attachment=True)


@app.route('/file-downloads/')
def file_downloads():

    return render_template("imprint.html")

@app.route("/graficos")
def graficos():

    cur_GraficoTotal = conexion.cursor()

    select_GraficoTotal = (""" select count(*) from reporte where fecha > sysdate - 7 AND fecha <= sysdate """)

    cur_GraficoTotal.execute(select_GraficoTotal)
    GraficoTotal = cur_GraficoTotal.fetchall()
    GraficoTotalf = [item[0] for item in GraficoTotal]
    

   


    
    result = ''.join(str(item) for item in GraficoTotalf)
 
    numero = int(result)
 

    #https://roytuts.com/google-pie-chart-using-python-flask/


    data = {'Task' : 'Hours per Day', 'Work' : numero , 'Eat' : 2, 'Commute' : 2, 'Watching TV' : 2, 'Sleeping' : 7}
    #print(data)
    
    return render_template("graficos.html",data=data,numero=numero)









    conexion.close()


if __name__ == "__main__":
    app.run(debug=True)  #Run app

