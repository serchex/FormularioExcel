import pandas as pd
import sys
import os

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView
)

df = pd.DataFrame(columns=['Nombre', 'Apellido', 'Profesion', 'Sueldo', 'Sueldo Neto'])

def crearApp():
    app = QApplication(sys.argv)
    ventana = QWidget()
    ventana.setWindowTitle('Registro de Empleados')
    ventana.resize(800,600)
    layout_principal=QVBoxLayout()
    layout_principal.addSpacing(30)
    layout_formulario=QHBoxLayout()
    
    nombre_input=QLineEdit()
    nombre_input.setFixedHeight(40)
    nombre_input.setPlaceholderText("Nombre")
    layout_formulario.addWidget(nombre_input)
    
    apellido_input=QLineEdit()
    apellido_input.setFixedHeight(40)
    apellido_input.setPlaceholderText("Apellido")
    layout_formulario.addWidget(apellido_input)
    
    profesion_input=QLineEdit()
    profesion_input.setFixedHeight(40)
    profesion_input.setPlaceholderText("Profesion")
    layout_formulario.addWidget(profesion_input)

    sueldo_input=QLineEdit()
    sueldo_input.setFixedHeight(40)
    sueldo_input.setPlaceholderText("Sueldo")
    layout_formulario.addWidget(sueldo_input)
    
    layout_principal.addLayout(layout_formulario)
    
    layout_botones=QHBoxLayout()
    boton_agregar=QPushButton("Agregar")
    #boton_agregar.setFixedHeight(40)
    layout_botones.addWidget(boton_agregar)

    boton_exportar=QPushButton("Exportar")
    layout_botones.addWidget(boton_exportar)
     
    layout_principal.addLayout(layout_botones)
    
    tabla=QTableWidget()
    tabla.setColumnCount(5)
    tabla.setHorizontalHeaderLabels(
        ["Nombre","Apellidos","Profesion","Sueldo","Sueldo Neto"]
    )
    tabla.horizontalHeader().setStretchLastSection(True)
    tabla.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
    layout_principal.addWidget(tabla)
    
    ventana.setLayout(layout_principal)
    ventana.show()
    
    def AgregarDatos():
        nombre=nombre_input.text()
        apellido=apellido_input.text()
        profesion=profesion_input.text()
        sueldo=float(sueldo_input.text())
        
        if sueldo > 0 and sueldo <= 100:
            sueldo_neto = sueldo*.80
        elif sueldo > 1000 and sueldo <= 4000:
            sueldo_neto = sueldo*.75
        else:
            sueldo_neto=sueldo*.65
            
        nueva_fila = {
            "Nombre":nombre,
            "Apellido":apellido,
            "Profesion":profesion,
            "Sueldo":sueldo,
            "Sueldo Neto":sueldo_neto
        }
        
        global df
        df.loc[len(df)] = nueva_fila
        fila = tabla.rowCount()
        tabla.insertRow(fila)
        tabla.setItem(fila,0,QTableWidgetItem(nombre))
        tabla.setItem(fila,1,QTableWidgetItem(apellido))
        tabla.setItem(fila,2,QTableWidgetItem(profesion))
        tabla.setItem(fila,3,QTableWidgetItem(f"{sueldo:.2f}"))            
        tabla.setItem(fila,4,QTableWidgetItem(f"{sueldo_neto:.2f}"))         
           
        tabla.resizeRowsToContents()
        nombre_input.clear()
        apellido_input.clear()
        profesion_input.clear()
        sueldo_input.clear()

    def ExportarExcel():
        if os.path.exists('dataEmpleados.xlsx'):
            os.remove('dataEmpleados.xlsx')
        df.to_excel('dataEmpleados.xlsx', index=False, engine='openpyxl')
    
    boton_agregar.clicked.connect(AgregarDatos)
    boton_exportar.clicked.connect(ExportarExcel)
    sys.exit(app.exec())
crearApp()