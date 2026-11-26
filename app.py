import tkinter as tk
import sqlite3
import hashlib
import openpyxl
from reportlab.pdfgen import canvas
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime, timedelta


ventana = tk.Tk()
ventana.title("AUTOCLOUD PARKING")
ventana.geometry("300x220")

def abrir_principal():
    principal = tk.Tk()
    principal.title("AUTOCLOUD PARKING V1.0")
    principal.geometry("800x500")
    contenido = tk.Frame(principal)
    contenido.pack(side="right", fill="both", expand=True)
    def limpiar():
        for cosa in contenido.winfo_children():
            cosa.destroy()
    def pantalla_ingreso():
        limpiar()
        tk.Label(contenido, text="REGISTRO DE VEHICULO",font=("Arial", 14, "bold")).pack(fill="x", pady=10)
        tk.Label(contenido, text="PLACA").pack(fill="x", padx=10, pady=5)
        caja_placa = tk.Entry(contenido)
        caja_placa.pack()
        tk.Label(contenido, text="TIPO DE VEHICULO").pack(fill="x",pady=5)
        caja_tipo= ttk.Combobox(contenido, values=["Carro", "Moto", "Bicicleta"], state="readonly")
        caja_tipo.pack()
        def registrar():
            placa = caja_placa.get().upper().strip()
            tipo = caja_tipo.get() 
            if placa == "" or tipo =="":
                messagebox.showerror("Error", "Por favor complete todos los campos") 
                return
            conection = sqlite3.connect("autocloudparking.db")
            dentro = conection.execute("SELECT * FROM registro WHERE placa=? AND salida IS NULL", (placa,)).fetchone()
            if dentro:
                messagebox.showerror("Error", "El vehiculo ya se encuentra registrado") 
            else:
                conection.execute("INSERT OR IGNORE INTO vehiculo values(?,?)", (placa,tipo))
                conection.execute("INSERT INTO registro (placa, entrada) VALUES (?, ?)", (placa, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
                conection.commit()
                messagebox.showinfo("Exito", "Registro exitoso")
                caja_placa.delete(0, "end")
            conection.close()
        tk.Button(contenido, text="REGISTAR",command=registrar, bg="yellow").pack(pady=10)
    def limpiar():
        for cosa in contenido.winfo_children():
            cosa.destroy()
    def pantalla_buscar_vehiculo():
        limpiar()
        tk.Label(contenido, text="BUSCAR VEHICULO",font=("Arial", 14, "bold")).pack(pady=10)
        tk.Label(contenido,text="PLACA").pack()
        caja_placa = tk.Entry(contenido)
        caja_placa.pack()
        resultado = tk.Label(contenido, text="", justify="left")
        resultado.pack(pady=15)
        def buscar():
            Placa = caja_placa.get().upper().strip()
            conecion = sqlite3.connect("autocloudparking.db") 
            vehiculo = conecion.execute("SELECT * FROM vehiculo WHERE placa=?", (Placa,)).fetchone()
            visitas = conecion.execute("SELECT entrada, salida FROM registro WHERE placa=?", (Placa,)).fetchall ()
            conecion.close()
            if vehiculo is None:
                messagebox.showerror("Error", "Vehiculo no encontrado")
                return
            texto = f"\nPLACA:{vehiculo[0]}\nTIPO: {vehiculo[1]}\n\nvisitas:"
            for v in visitas:
                texto = texto + f"\nentrada: {v[0]} salida: {v[1] or 'En el Parqueadero'} "
            resultado.config(text=texto)

        tk.Button(contenido, text="BUSCAR",command=buscar, bg="yellow").pack(pady=10)
    def limpiar():
        for cosa in contenido.winfo_children():
            cosa.destroy()   
    def pantalla_calcular_tarifa():
        limpiar()
        tk.Label(contenido, text="CALCULAR TARIFA", font=("Arial", 14, "bold")).pack(pady=10)
        caja_placa= tk.Entry(contenido)
        caja_placa.pack()
        resultado = tk.Label(contenido, text="", font=("Arial",12))
        resultado.pack(pady=15)
        datos = {}

        def calcular():
            placa = caja_placa.get().upper().strip()
            conecion= sqlite3.connect("autocloudparking.db")
            fila = conecion.execute("SELECT registro.entrada, tarifa.valor_hora FROM registro JOIN vehiculo ON registro.placa = vehiculo.placa JOIN tarifa ON vehiculo.tipo = tarifa.tipo  WHERE registro.placa=? AND registro.salida IS NULL", (placa,)).fetchone()  
            conecion.close()
            if fila is None:
                messagebox.showerror("Error", "vehiculo no encontrado")
            else:
                entrada = datetime.strptime(fila[0], "%d/%m/%Y %H:%M:%S")
                segundos = (datetime.now() - entrada).total_seconds() 
                horas = int(segundos // 3600) + 1
                valor= horas * fila[1]
                datos["placa"] = placa
                datos["valor"] = valor
                resultado.config(text=f"Horas: {horas} Valor a pagar: $ {valor}")
        def confirmar():
                if "placa" not in datos:
                    messagebox.showerror("Error", "Calcular tarifa")
                    return
                conecion = sqlite3.connect("autocloudparking.db")
                conecion.execute("UPDATE registro SET salida=?, valor=? WHERE placa=? AND salida is NULL", (datetime.now().strftime("%d/%m/Y %H:%M:%S"), datos["valor"], datos["placa"] ))
                conecion.commit()
                conecion.close()
                messagebox.showinfo("Confirmar", "Confirmacion exitosa")
                datos.clear()
                resultado.config(text="")
                caja_placa.delete(0, "end")       
        tk.Button(contenido, text="CALCULAR",command=calcular, bg="yellow").pack(pady=10)
        tk.Button(contenido, text="CONFIRMAR", command=confirmar, bg="green").pack(pady=10)
       
    def limpiar():
        for cosa in contenido.winfo_children():
            cosa.destroy()    
    def pantalla_generar_reporte():
        limpiar()
        tk.Label(contenido, text="GENERAR REPORTE", font=("Arial", 14,"bold")).pack(pady=10)
        
        tk.Label(contenido, text="PERIODO", font=("Arial",12, "bold")).pack(pady=10)
        tk.Label(contenido, text="PERIODO").pack
        caja_periodo = ttk.Combobox(contenido, values=["Diario", "Semanal", "Mensual", "Anual"], state="readonly")
        caja_periodo.pack()
        tk.Label(contenido, text="TIPO VEHICULO", font=("Arial",12, "bold")).pack(pady=10)
        caja_tipo = ttk.Combobox(contenido, values=["Todos", "Carro", "Moto", "Bicilceta"], state="readonly")
        caja_tipo.pack()
        resultado = tk.Label(contenido, text="", font=("Arial", 12, "bold" ),justify="left")
        resultado.pack(pady=15)
        def consultar():
            periodo = caja_periodo.get()
            tipo = caja_tipo.get()
            if  periodo == "" or tipo =="":
             messagebox.showerror("Error", "Selecone el periodo y tipo de vehiculo")
             return None
            hoy = datetime.now()
            if periodo == "Diario":
                desde = hoy 
            elif periodo == "Semanal":
                desde = hoy - timedelta(days=6)
            elif periodo == "Mensual":
                desde = hoy.replace(day=1)
            else: desde = hoy.replace(month=1, day=1) 
            sql = "SELECT registro.placa, vehiculo.tipo, registro.entrada, registro.salida, registro.valor FROM registro JOIN vehiculo ON registro.placa = vehiculo.placa WHERE registro.salida >= ?"
            valores = [desde.strftime("%d/%m/%y")]
            if tipo != "Todos":
                sql = sql + " AND vehiculo.tipo = ?"
                valores.append(tipo)
            conecion = sqlite3.connect("autocloudparking.db")
            filas= conecion.execute(sql,valores).fetchall()
            conecion.close() 
            return filas
        def generar():
            filas = consultar()
            if filas is None:
                return
            total =sum(f[4] for  f in filas) 
            resultado.config(text=f"Vehiculos atendidos:{len(filas)}\nTotal recaudado ${total}")

        def excel():
                filas = consultar() 
                if filas is None:
                    return
                archivo = filedialog.asksaveasfilename(defaultextension=".xslx", filetypes=[("Excel", "*.xslx")])
                if archivo !="":
                    libro = openpyxl.Workbook()
                    hoja = libro.active
                    hoja.append(["Placa", "Tipo", "Entrada", "Salida", "Valor"])
                    for fila in filas:   
                        hoja.append(fila)
                    hoja.append(["","","","TOTAL", sum(f[4] for f in filas)])
                    libro.save(archivo)
                    messagebox.showinfo("Excel", "Reporte guardado")
        def pdf():
            filas = consultar()
            if filas is None: 
                return
            archivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if archivo =="":
                return
            titulos = ["Placa", "Tipo", "Entrada", "Salida", "Valor"]
            columnas =[50,120,190,320,450]
            hoja = canvas.Canvas(archivo)
            hoja.setFont("Helvetica-Bold", 14)
            hoja.drawString(50, 800, f"Reporte{caja_periodo.get()} - {caja_tipo.get()}")
            hoja.setFont("Helvetica-Bold",10)
            for i in range(5):
                hoja.drawString(columnas[i], 770, titulos[i])
            hoja.setFont("Helvetica", 10)    
            y = 750
            for f in filas:
                for i in range(5):
                    hoja.drawString(columnas[i], y, str(f[i]))
                y = y -18
                if y <60:
                    hoja.showPage()
                    hoja.setFont("Helvetica", 10)
                    y = 800 
            hoja.drawString(50, y -10, f"Total vehiculos: {len(filas)} Total tarifas: ${sum(f[4] for f in filas)}")
            hoja.save()
            messagebox.showinfo("PDF", "Reporte guardado")
              
        tk.Button(contenido, text="GENERAR",command=generar, bg ="yellow").pack(pady=10)
        tk.Button(contenido,text="DESCARGAR formato EXCEL",command=excel, bg="green").pack(pady=10)
        tk.Button(contenido, text="DESCARGAR formato PDF", command=pdf, bg="red").pack(pady=10)

        
    menu = tk.Frame(principal, bg="Blue")
    menu.pack(side="left", fill="y")
    tk.Label(menu, text="MENU PRINCIPAL",bg="Blue", fg="White").pack(pady=10)
    tk.Button(menu, text="REGISTRO DE VEHICULO",command=pantalla_ingreso).pack(fill="x", padx=10, pady=5)
    tk.Button(menu, text="BUSCAR VEHICULO",command=pantalla_buscar_vehiculo).pack(fill="x", padx=10,pady=5)
    tk.Button(menu, text="CALCULAR TARIFA",command=pantalla_calcular_tarifa ).pack(fill="x",padx=10,pady=5)
    tk.Button(menu, text="GENERAR REPORTE",command=pantalla_generar_reporte).pack(fill="x",padx=10,pady=5)
    tk.Button(menu, text="CERRAR SESION",command=principal.destroy,bg="red",fg="black").pack(fill="x",padx=10,pady=5)
    principal.mainloop()
def registar_usuario(evento=None):
    nueva = tk.Toplevel(ventana)
    nueva.title("REGISTAR USUARIO")
    nueva.geometry("300x280")

    tk.Label(nueva, text="USUARIO NUEVO").pack(pady=5) 
    caja_nuevo = tk.Entry(nueva)
    caja_nuevo.pack()
    tk.Label(nueva, text="CONTRASEÑA").pack(pady=5)
    caja_clave = tk.Entry(nueva, show="*")
    caja_clave.pack()
    tk.Label(nueva, text="CONFIRMAR CONTRASEÑA").pack(pady=5)
    caja_clave1 =tk.Entry(nueva, show="*")
    caja_clave1.pack() 

    def guardar():
        usuario = caja_nuevo.get().strip()
        clave   = caja_clave.get()
        if usuario =="" or clave =="":
            messagebox.showerror("Error", "Registre todos los campos")
            return
        if clave != caja_clave1.get():
            messagebox.showerror("Error", "Las contraseñas no coinciden")
            return
        conecion = sqlite3.connect("autocloudparking.db")
        existe = conecion.execute("SELECT * FROM usuario  Where usuario=?", (usuario,)).fetchone()
        if existe:
            messagebox.showerror("Error", "El usario ya existe")
        else:
            conecion.execute("INSERT INTO usuario VALUES (?,?)", (usuario, hashlib.sha256(clave.encode()).hexdigest()))
            conecion.commit()
            messagebox.showinfo("Usuario", "Usuario credo con exitos")
            nueva.destroy()
        conecion.close()
    tk.Button(nueva, text="GUARDAR", command=guardar, bg="blue" ).pack(pady=15)
         
def ingresar():
    clave = hashlib.sha256(caja_contraseña.get().encode()).hexdigest()
    conecion = sqlite3.connect("autocloudparking.db")
    ficha = conecion.execute("SELECT * FROM usuario WHERE usuario=? AND clave =?",(caja_usuario.get(), clave)).fetchone()
    conecion.close()
    if ficha:
        messagebox.showinfo("Bienvenido", "INGRESO EXITOSO")
        ventana.destroy()
        abrir_principal()
    else:
        messagebox.showerror("Error", "Usuario o contraseña incorrectos")
tk.Label(ventana, text="USUARIO").pack (pady=5)
caja_usuario = tk.Entry(ventana)
caja_usuario.pack()

tk.Label(ventana, text="CONTRASEÑA").pack (pady=5)
caja_contraseña = tk.Entry(ventana, show="*")
caja_contraseña.pack()

tk.Button(ventana, text="INGRESAR",command=ingresar, bg="yellow",fg="black").pack(pady=15)

enlace = tk.Label(ventana, text="¿No tiesnes cuenta? Registrar usuario", fg="blue", cursor="hand2", font=("Arial", 9, "underline"))
enlace.pack(pady=5)
enlace.bind("<Button-1>", registar_usuario)

ventana.mainloop()