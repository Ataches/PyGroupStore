from tkinter import *
from random import random
from tkinter import messagebox
# Lectura
from xlrd import open_workbook
# Escritura
from tempfile import TemporaryFile
from xlwt import Workbook, easyxf
# Modificacion de existente
from xlutils.copy import copy
# manejo de fechas
import datetime

archivo = "data/datos.xls"

productos = []
vTotal = 0


def alerta(mensaje):
    messagebox.showinfo("Alerta", message=mensaje)


# Acciones de los botones
def obtenerDato():
    global vTotal
    guardarFactura(idFactura, entradaN.get(), entradaD.get(), entradaT.get(), listaCity.get(ACTIVE), str(hoy),
                   productos, vTotal)
    ventana.destroy()


def obtenerArticulo():
    global vTotal
    vUni = int(vUnitario.get())
    cant = int(cantidad.get())
    productos.append([cant, referencia.get(), descp.get(), vUni, vUni * cant])
    vTotal += (vUni * cant)
    mensaje = "Se guardo el siguiente articulo: \n\nCantidad: ", cant, "\nreferencia: ", referencia.get(), "\ndescripcion ", descp.get(), "\nvalor unitario: ", vUni,
    mensaje += "\nvalor total: ", vUni * cant, "\n\n\t Total factura: ", vTotal
    alerta(mensaje)
    print(productos)


# Archivos
# Metodo de guardar en archivo existente
def guardarFactura(idFact, nombre, direccion, tel, ciudad, fecha, prod, total):
    wb = open_workbook(archivo)
    s = wb.sheet_by_name("Facturas")
    fila = int(s.cell(0, 0).value)  # Indice de la 1a linea vacia

    rb = open_workbook(archivo)
    wb = copy(rb)
    s = wb.get_sheet(0)
    s.write(fila, 0, idFact)
    s.write(fila, 1, nombre)
    s.write(fila, 2, direccion)
    s.write(fila, 3, tel)
    s.write(fila, 4, ciudad)
    s.write(fila, 5, fecha)
    s.write(fila, 6, total)
    cantProd = len(prod)
    s.write(fila, 7, cantProd)
    for i in range(cantProd):
        item = prod[i]
        s.write(fila, (8 + (i * 5)), item[0])  # cant)
        s.write(fila, (9 + (i * 5)), item[1])  # referencia)
        s.write(fila, (10 + (i * 5)), item[2])  # descripcion)
        s.write(fila, (11 + (i * 5)), item[3])  # vUnitario)
        s.write(fila, (12 + (i * 5)), item[4])  # vTotal)
    s.write(0, 0, (fila + 1))
    wb.save(archivo)

    msj = "\n\t\tID: ", idFact
    msj += "\n\tNombre: ", nombre
    msj += "\n\tDireccion: ", direccion
    msj += "\n\tTelefono: ", tel
    msj += "\n\tCiudad: ", ciudad
    msj += "\n\tFecha: ", fecha
    msj += "\n\nProductos: ", prod, "\n\n"
    msj += "Total: ", total
    alerta(msj)
    msj = "Se guardo con exito su factura en ", archivo
    alerta(msj)

    # Metodo buscar factura por ID


def mostrarFactura():
    mostrarFactura = False
    try:
        idBuscada = int(idBsqda.get())

        wb = open_workbook(archivo)
        s = wb.sheet_by_name("Facturas")
        infoFactura = []
        productos = []

        for row in range(2, s.nrows):
            if (idBuscada == int(s.cell(row, 0).value)):
                mostrarFactura = True
                break
        if(not mostrarFactura):
            raise ValueError
    except:
        alerta("ID No identificada en el archivo de datos.xls")
        return
    if (mostrarFactura):
        for col in range(0, 8):
            infoFactura.append(s.cell(row, col).value)
        for col in range(8, 8 + (5 * int(s.cell(row, 7).value))):
            productos.append(s.cell(row, col).value)
        infoFactura.append(productos)
        generarRecibo(infoFactura)


def generarRecibo(infoFactura):
    book = Workbook()
    s = book.add_sheet('Recibo generado')
    s.write(3, 3, "PYGROUP STORE", easyxf(
        'font: name Arial, height 300, bold True;'
    ))
    s.write(1, 4, "FACTURA ID:", easyxf(
        'font: name Arial, bold True, height 220;'
        'borders: left thick, top thick, bottom thick;'
    ))
    s.write(1, 5, infoFactura[0], easyxf(
        'font: name Arial, height 220;'
        'borders: right thick, top thick, bottom thick;'
    ))
    s.write(6, 2, "Nombre:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: left thick, top thick;'
    ))
    s.write(6, 3, infoFactura[1], easyxf(
        'font: name Arial, height 200;'
        'borders: right thick, top thick;'
    ))
    s.write(7, 2, "Dirección:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: left thick;'
    ))
    s.write(7, 3, infoFactura[2], easyxf(
        'font: name Arial, height 200;'
        'borders: right thick;'
    ))
    s.write(8, 2, "Telefono:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: left thick, bottom thick;'
        'alignment: horizontal left;'
    ))
    s.write(8, 3, infoFactura[3], easyxf(
        'font: name Arial, height 200;'
        'borders: right thick, bottom thick;'
    ))
    s.write(10, 2, "Ciudad:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: left thick, top thick;'
    ))
    s.write(10, 3, infoFactura[4], easyxf(
        'font: name Arial, height 200;'
        'borders: right thick, top thick;'
    ))
    s.write(11, 2, "Fecha:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: left thick, bottom thick;'
    ))
    s.write(11, 3, infoFactura[5], easyxf(
        'font: name Arial, height 200;'
        'borders: right thick, bottom thick;'
    ))
    mensaje = "\n\tID: ", infoFactura[0], "\n\tNombre: ", infoFactura[1], "\n\tDirección: ", infoFactura[
        2], "\n\tTelefono: ", infoFactura[3], "\n\n\tCiudad: ", infoFactura[4], "\n\tFecha: ", infoFactura[
                  5], "\n\nProductos facturados:"
    cantProd = len(infoFactura[8])
    s.col(0).width = 1000
    s.col(2).width = 3800
    s.col(3).width = 10000
    s.col(4).width = 3800
    s.col(5).width = 3800
    s.write(14, 1, "CANT:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: bottom thick, left thick, top thick;'
        'alignment: horizontal center;'
    ))
    s.write(14, 2, "REFERENCIA:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: bottom thick, top thick;'
        'alignment: horizontal center;'
    ))
    s.write(14, 3, "DESCRIPCIÓN:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: bottom thick, top thick;'
        'alignment: horizontal center;'
    ))
    s.write(14, 4, "V UNITARIO:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: bottom thick, top thick;'
        'alignment: horizontal center;'
    ))
    s.write(14, 5, "V TOTAL:", easyxf(
        'font: name Arial, bold True, height 200;'
        'borders: bottom thick,right thick, top thick;'
        'alignment: horizontal center;'
    ))
    for i in range(0, cantProd, 5):
        item = infoFactura[8]
        s.write(int(15 + i / 5), 1, item[i], easyxf(
            'alignment: horizontal center;'
        ))  # cant)
        s.write(int(15 + i / 5), 2, item[i + 1], easyxf(
            'alignment: horizontal center;'
        ))  # referencia)
        s.write(int(15 + i / 5), 3, item[i + 2], easyxf(
            'alignment: horizontal center;'
        ))  # descripcion)
        s.write(int(15 + i / 5), 4, item[i + 3])  # vUnitario)
        s.write(int(15 + i / 5), 5, item[i + 4])  # vTotal)
        mensaje += "\n\nCantidad: ", item[i], " referencia: ", item[i + 1], " descripcion ", item[
            i + 2], " valor unitario: ", item[i + 3],
        mensaje += "valor total: ", item[i + 4]
    s.write(int(cantProd / 5) + 19, 4, "TOTAL: ", easyxf(
        'font: name Arial, bold True, height 260;'
        'borders: left thick, top thick, bottom thick;'
    ))
    s.write(int(cantProd / 5) + 19, 5, infoFactura[6], easyxf(
        'font: name Arial, height 220;'
        'borders: right thick, top thick, bottom thick;'
    ))
    mensaje += "\n\n\tTOTAL: ", infoFactura[6]
    alerta(mensaje)
    book.save('data/Recibo.xls')
    book.save(TemporaryFile())
    alerta("Se genero su recibo con exito en el archivo Recibo.xls")


# INICIO DE VENTANA

indice = int(3)

ventana = Tk()
ventana.geometry("900x660")  # 750-800
ventana.title('Empresa ')

# Fecha y random id factura
hoy = datetime.datetime.now()
fecha = Label(ventana, text=hoy, fg="green").place(x=20, y=20)
ciudad = "Bogota"

idFactura = str(int(random() * 10000 + 1))
idF = Label(ventana, text=("ID: " + idFactura), fg="blue").place(x=20, y=40)

# fields y etiquetas

# Campos de identificación del cliente
etiqueta_name = Label(ventana, text="Nombre", font=("Arial Black", 12)).place(x=20, y=250)
entradaN = StringVar()
text_field_Name = Entry(ventana, width=30, textvariable=entradaN).place(x=20, y=280)

etiqueta_tel = Label(ventana, text="Telefono", font=("Arial Black", 12)).place(x=230, y=250)
entradaT = StringVar()
text_field_Tel = Entry(ventana, width=30, textvariable=entradaT).place(x=230, y=280)

etiqueta_dir = Label(ventana, text="Direccion", font=("Arial Black", 12)).place(x=460, y=250)
entradaD = StringVar()
text_field_Dir = Entry(ventana, width=30, textvariable=entradaD).place(x=460, y=280)

etiqueta_city = Label(ventana, text="Ciudad", font=("Arial Black", 12)).place(x=680, y=250)
listaCity = Listbox(ventana, width=30, height=1)
listaCity.insert(0, "Bogota")
listaCity.insert(1, "Medellin")
listaCity.insert(2, "Neiva")
listaCity.insert(3, "Huila")
listaCity.insert(4, "Cali")
listaCity.place(x=680, y=280)

# Productos a insertar
etiqueta_namep = Label(ventana, text="Cantidad", font=("Agency FB", 14)).place(x=150, y=390)
cantidad = StringVar()
text_field_cant = Entry(ventana, width=8, textvariable=cantidad).place(x=150, y=420)

etiqueta_tel = Label(ventana, text="Referencia", font=("Agency FB", 14)).place(x=230, y=390)
referencia = StringVar()
text_field_ref = Entry(ventana, width=16, textvariable=referencia).place(x=230, y=420)

etiqueta_dir = Label(ventana, text="Descripción", font=("Agency FB", 14)).place(x=360, y=390)
descp = StringVar()
text_field_descp = Entry(ventana, width=50, textvariable=descp).place(x=360, y=420)

etiqueta_vu = Label(ventana, text="Valor unitario", font=("Agency FB", 14)).place(x=700, y=390)
vUnitario = StringVar()
text_field_vUni = Entry(ventana, width=16, textvariable=vUnitario).place(x=700, y=420)

etiqueta_bf = Label(ventana, text="Buscar factura por ID y generar su recibo: ").place(x=20, y=520)
idBsqda = StringVar()
text_field_vUni = Entry(ventana, width=16, textvariable=idBsqda).place(x=45, y=570)

# Botones
# Guardar productos
btnGenerarRecibo = Button(ventana, text="Guardar producto", command=obtenerArticulo, height=1, width=20,
                          activebackground="red").place(x=740, y=520)
# Buscar productos
btnBuscarFact = Button(ventana, text="Buscar ID", command=mostrarFactura, height=1, width=20).place(x=20, y=600)
# Guardar factura
btnGenerarRecibo = Button(ventana, text="Guardar factura", command=obtenerDato, height=2, width=20,
                          activebackground="#82FA58").place(x=740, y=580)

etiqueta_div = Label(ventana,
                     text="______________________________________________________________________________________________________________________________________",
                     fg="grey").place(x=145, y=340)

# fin fields

# Carga de imagenes en la interfaz
logo = PhotoImage(file="images/logoad.gif")
etiquetaLogo = Label(ventana, image=logo).place(x=200, y=20)

logousr = PhotoImage(file="images/logousr.gif")
etiquetaLogo = Label(ventana, image=logousr).place(x=0, y=200)

logopr = PhotoImage(file="images/logopr.gif")
etiquetaLogo = Label(ventana, image=logopr).place(x=0, y=340)

logoud = PhotoImage(file="images/logoudmini.png")
etiquetaLogo = Label(ventana, image=logoud).place(x=20, y=60)

if __name__ == '__main__':
    ventana.mainloop()
