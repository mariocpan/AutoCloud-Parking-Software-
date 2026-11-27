import tkinter as tk
import sqlite3
import hashlib
import openpyxl
import db
import ctypes

from reportlab.pdfgen import canvas
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime, timedelta

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("autocloudparking")
ventana = tk.Tk()
ventana.iconbitmap("icono.ico")
ventana.title("AUTOCLOUD PARKING")
ventana.geometry("300x520")

def abrir_principal():
    principal = tk.Tk()
    principal.title("AUTOCLOUD PARKING V1.0")
    principal.iconbitmap("icono.ico")
    principal.geometry("800x500")
    contenido = tk.Frame(principal, bg="#DEEAF6")
    contenido.pack(side="right", fill="both", expand=True)
    def limpiar():
        for cosa in contenido.winfo_children():
            cosa.destroy()
    def pantalla_ingreso():
        limpiar()
        tk.Label(contenido, text="REGISTRO DE VEHICULO",font=("Arial", 16, "bold"), fg="#1B5BA5", bg="#DEEAF6").pack(fill="x", pady=28)
        tk.Label(contenido, text="PLACA O DOCUMENTO",font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
        caja_placa = tk.Entry(contenido,width=25)
        caja_placa.pack()
        tk.Label(contenido, text="TIPO DE VEHICULO",font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x",pady=5)
        caja_tipo= ttk.Combobox(contenido, values=["Carro", "Moto", "Bicicleta"],width=22, state="readonly")
        caja_tipo.pack(pady=5)
        def registrar():
            placa = caja_placa.get().upper().strip()
            tipo = caja_tipo.get() 
            if placa == "" or tipo =="":
                messagebox.showerror("Error", "Por favor, complete todos los campos") 
                return
            if tipo =="Bicicleta":
                if not placa.isdigit() or len(placa) < 6 or len(placa) > 10:
                    messagebox.showerror("Error", "Documento invalido: solo numeros, de 6 a 10 digitos")    
                    return
            elif not placa.isalnum() or len(placa) < 5 or len(placa)>7:
                messagebox.showerror("Error", "Placa inválida: debe contener entre 5 y 7 caracteres alfanuméricos, sin espacios ni guiones.")
                return
            conection = sqlite3.connect("autocloudparking.db")
            dentro = conection.execute("SELECT * FROM registro WHERE placa=? AND salida IS NULL", (placa,)).fetchone()
            if dentro:
                messagebox.showerror("Error", "El vehiculo ya se encuentra registrado") 
            else:
                conection.execute("INSERT OR IGNORE INTO vehiculo values(?,?)", (placa,tipo))
                conection.execute("INSERT INTO registro (placa, entrada) VALUES (?, ?)", (placa, datetime.now().strftime("%Y/%m/%d %H:%M:%S")))
                conection.commit()
                messagebox.showinfo("Exito", "Registro exitoso")
                caja_placa.delete(0, "end")
            conection.close()
        tk.Button(contenido, text="REGISTRAR",command=registrar,bg="#F8B810",fg="#12345A", font=("Arial", 10, "bold"), width=15).pack(pady=5)
        
    def pantalla_buscar_vehiculo():
        limpiar()
        tk.Label(contenido, text="BUSCAR VEHICULO",fg="#1B5BA5", bg="#DEEAF6", font=("Arial", 16, "bold")).pack(fill="x",pady=20)
        tk.Label(contenido,text="PLACA O DOCUMENTO",font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
        caja_placa = tk.Entry(contenido, width=25)
        caja_placa.pack(pady=5)
        resultado = tk.Label(contenido, text="",font=("Arial", 12, "bold"), fg="#12345A", bg="#DEEAF6")
        resultado.pack(pady=15)
        tabla = ttk.Treeview(contenido, columns=("entrada", "salida"), show="headings", height=6)
        tabla.heading("entrada", text="ENTRADA")
        tabla.heading("salida", text="SALIDA")
        tabla.column("entrada", width=150, anchor="center")
        tabla.column("salida", width=150, anchor="center")
        tabla.pack(pady=10)
        def buscar():
            Placa = caja_placa.get().upper().strip()
            conecion = sqlite3.connect("autocloudparking.db")
            vehiculo = conecion.execute("SELECT * FROM vehiculo WHERE placa=?", (Placa,)).fetchone()
            visitas = conecion.execute("SELECT entrada, salida FROM registro WHERE placa=?", (Placa,)).fetchall()
            conecion.close()
            if vehiculo is None:
                messagebox.showerror("Error", "Vehiculo no encontrado")
                return
            tabla.delete(*tabla.get_children())
            resultado.config(text=f"PLACA: {vehiculo[0]}     TIPO: {vehiculo[1]}")
            for v in visitas:
                tabla.insert("", "end", values=(v[0], v[1] or "En el parqueadero"))

        tk.Button(contenido, text="BUSCAR",command=buscar, bg="#F8B810",fg="#12345A", font=("Arial", 10, "bold"), width=15).pack(pady=20)
    
    def pantalla_calcular_tarifa():
        limpiar()
        tk.Label(contenido, text="CALCULAR TARIFA", font=("Arial", 16, "bold"), fg="#1B5BA5", bg="#DEEAF6").pack(fill="x",pady=10)
        tk.Label(contenido, text="PLACA O DOCUMENTO", font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
        caja_placa= tk.Entry(contenido,width=25)
        caja_placa.pack()
        resultado = tk.Label(contenido, text="", font=("Arial",12),fg="#12345A", bg="#DEEAF6")
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
                entrada = datetime.strptime(fila[0], "%Y/%m/%d %H:%M:%S")
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
                conecion.execute("UPDATE registro SET salida=?, valor=? WHERE placa=? AND salida is NULL", (datetime.now().strftime("%Y/%m/%d %H:%M:%S"), datos["valor"], datos["placa"] ))
                conecion.commit()
                conecion.close()
                messagebox.showinfo("Confirmar", "Confirmacion exitosa")
                datos.clear()
                resultado.config(text="")
                caja_placa.delete(0, "end")       
        tk.Button(contenido, text="CALCULAR",command=calcular,bg="#F8B810", fg="#12345A",font=("Arial", 10, "bold"), width=15).pack(pady=10)
        tk.Button(contenido, text="CONFIRMAR", command=confirmar, bg="#1B5BA5", fg="white",font=("Arial", 10, "bold"), width=15).pack(pady=10)
       
    def pantalla_generar_reporte():
        limpiar()
        tk.Label(contenido, text="GENERAR REPORTE", font=("Arial", 16, "bold"), fg="#1B5BA5", bg="#DEEAF6").pack(fill="x", pady=28)
        tk.Label(contenido, text="PERIODO", font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
        caja_periodo = ttk.Combobox(contenido, values=["Diario", "Semanal", "Mensual", "Anual"], width=22, state="readonly")
        caja_periodo.pack(pady=5)
        tk.Label(contenido, text="TIPO VEHICULO", font=("Arial",10, "bold"),fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
        caja_tipo = ttk.Combobox(contenido, values=["Todos", "Carro", "Moto", "Bicicleta"], width=22, state="readonly")
        caja_tipo.pack(pady=5)
        resultado = tk.Label(contenido, text="",bg="#DEEAF6", font=("Arial", 12, "bold" ),justify="left")
        resultado.pack(pady=15)
        def consultar():
            periodo = caja_periodo.get()
            tipo = caja_tipo.get()
            if  periodo == "" or tipo =="":
             messagebox.showerror("Error", "Seleccione el periodo y tipo de vehiculo")
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
            valores = [desde.strftime("%Y/%m/%d")]
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
                archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
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
            hoja.drawString(50, 800, f"Reporte {caja_periodo.get()} - {caja_tipo.get()}")
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
              
        tk.Button(contenido, text="GENERAR",command=generar,bg="#F8B810", fg="#12345A", font=("Arial", 10, "bold"), width=15).pack(pady=10)
        tk.Button(contenido,text="DESCARGAR formato EXCEL",command=excel,bg="#1B5BA5", fg="white", font=("Arial", 10, "bold"), width=25).pack(pady=5)
        tk.Button(contenido, text="DESCARGAR formato PDF", command=pdf,bg="#1B5BA5", fg="white", font=("Arial", 10, "bold"), width=25).pack(pady=5)

    
    menu = tk.Frame(principal, bg="#DEEAF6")
    menu.pack(side="left", fill="y")
    logo = tk.PhotoImage(file="logo.png", master=principal).subsample(3)
    tk.Label(menu, image=logo, bg="#DEEAF6").pack(pady=15)
    tk.Label(menu, text="MENU PRINCIPAL", bg="#DEEAF6", fg="#1B5BA5", font=("Arial", 11, "bold")).pack(pady=10)

    estilo = dict(bg="#F8B810", fg="#12345A", font=("Arial", 10, "bold"))
    tk.Button(menu, text="REGISTRO DE VEHICULO", command=pantalla_ingreso, **estilo).pack(fill="x", padx=10, pady=5)
    tk.Button(menu, text="BUSCAR VEHICULO", command=pantalla_buscar_vehiculo, **estilo).pack(fill="x", padx=10, pady=5)
    tk.Button(menu, text="CALCULAR TARIFA", command=pantalla_calcular_tarifa, **estilo).pack(fill="x", padx=10, pady=5)
    tk.Button(menu, text="GENERAR REPORTE", command=pantalla_generar_reporte, **estilo).pack(fill="x", padx=10, pady=5)
    tk.Button(menu, text="CERRAR SESION", command=principal.destroy, bg="#1B5BA5", fg="white", font=("Arial", 10, "bold")).pack(fill="x", padx=10, pady=5)
    pantalla_ingreso()
    principal.mainloop()
def registar_usuario(evento=None):
    nueva = tk.Toplevel(ventana)
    nueva.title("REGISTRAR USUARIO")
    nueva.iconbitmap("icono.ico")
    nueva.geometry("300x520")
    nueva.config(bg="#DEEAF6")

    tk.Label(nueva, text="USUARIO NUEVO",font=("Arial", 16, "bold"), fg="#1B5BA5", bg="#DEEAF6").pack(fill="x", pady=20)
    tk.Label(nueva, text="USUARIO", font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5) 
    caja_nuevo = tk.Entry(nueva, width=25)
    caja_nuevo.pack(pady=5)
    tk.Label(nueva, text="CONTRASEÑA",font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
    caja_clave = tk.Entry(nueva, width=25, show="*")
    caja_clave.pack(pady=5)
    tk.Label(nueva, text="CONFIRMAR CONTRASEÑA", font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
    caja_clave1 =tk.Entry(nueva, width=25, show="*")
    caja_clave1.pack(pady=5) 

    def guardar():
        usuario = caja_nuevo.get().strip()
        clave   = caja_clave.get()
        if usuario =="" or clave =="":
            messagebox.showerror("Error", "Registre todos los campos", parent=nueva)
            return
        if len(usuario) < 3 or " " in usuario:
            messagebox.showerror("Error", "El usuario debe tener minimo 3 caracteres y sin espacios", parent=nueva)
            return
        if len(clave) < 6:
            messagebox.showerror("Error","La contraseña debe tener minimo 6 caracteres", parent=nueva) 
            return
        if clave != caja_clave1.get():
            messagebox.showerror("Error", "Las contraseñas no coinciden",parent=nueva)
            return
        conecion = sqlite3.connect("autocloudparking.db")
        existe = conecion.execute("SELECT * FROM usuario  Where usuario=?", (usuario,)).fetchone()
        if existe:
            messagebox.showerror("Error", "El usuario ya existe",parent=nueva)
        else:
            conecion.execute("INSERT INTO usuario VALUES (?,?)", (usuario, hashlib.sha256(clave.encode()).hexdigest()))
            conecion.commit()
            messagebox.showinfo("Usuario", "Usuario creado con éxito",parent=nueva)
            nueva.destroy()
        conecion.close()
    tk.Button(nueva, text="GUARDAR", command=guardar,bg="#F8B810", fg="#12345A", font=("Arial", 10, "bold"), width=15).pack(pady=20)
         
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

logo = tk.PhotoImage(file="logo.png").subsample(2)
tk.Label(ventana, image=logo, bg="#DEEAF6").pack(pady=10)

ventana.configure(bg="#DEEAF6")
tk.Label(ventana, text="USUARIO", font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
caja_usuario = tk.Entry(ventana, width=25)
caja_usuario.pack(pady=5)

tk.Label(ventana, text="CONTRASEÑA",font=("Arial", 10, "bold"), fg="#12345A", bg="#DEEAF6").pack(fill="x", padx=10, pady=5)
caja_contraseña = tk.Entry(ventana, width=25, show="*")
caja_contraseña.pack(pady=5)

tk.Button(ventana, text="INGRESAR",command=ingresar,bg="#F8B810", fg="#12345A", font=("Arial", 10, "bold"), width=15).pack(pady=15)

enlace = tk.Label(ventana, text="¿No tienes cuenta? Registrar usuario",fg="#1B5BA5", bg="#DEEAF6", cursor="hand2", font=("Arial", 9, "underline"))
enlace.pack(pady=5)
enlace.bind("<Button-1>", registar_usuario)

ventana.mainloop()