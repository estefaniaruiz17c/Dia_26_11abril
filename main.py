# Dataframes con pandas utilizando archivos de excel
print("Dataframes con pandas utilizando archivos de excel")

# Importar las librerías necesarias para el ejercicio
import openpyxl
import pandas as pd

# Crearemos la variable para leer el excel 1, contiene números enteros solamente
num_1 = pd.read_excel("dataframe1.xlsx",sheet_name=0)
print(num_1,"\n")

# Esta variable recorre todo el Dataframe
dataf1 = pd.DataFrame(num_1)

# Procederemos a crear el archivo donde guardaremos las operaciones 

# Crear un archivo de Excel
opera_doc_1 = openpyxl.Workbook()

# Aquí asignamos una hoja de cálculo en blanco en el archivo creado en el paso anterior
hojacalculo = opera_doc_1.active

# Guardaremos lo que llevamos en el archivo con el nombre: 'opera1_excel_1.xlsx'
opera_doc_1.save('opera1_excel_1.xlsx')

# Ya creado, continuaremos diseñando algunas operaciones
hojacalculo['A1'] = ("Operaciones realizadas con la librería pandas de Pyhton")
hojacalculo['A3'] = ("Ejercicios parte 1")
hojacalculo['A4'] = 1
hojacalculo['A5'] = 2
hojacalculo['A6'] = 3
hojacalculo['A7'] = 4

# Con esto, haremos algunos ejercicios

# Ejercicio 1: shape - devuelve una tupla con el número de filas y columnas del DataFrame
hojacalculo['B4'] = ("Número de filas y columnas del DataFrame del archivo 'dataframe1.xlsx")
print("shape")

# Creación de la operación
shape1 = str(dataf1.shape) 
print(shape1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C4'] = shape1


# Ejercicio 2: multiplicar datos de la columna c por 10
hojacalculo['B5'] = ("Multiplicar por 10 los datos de la columna 'c' de 'dataframe1.xlsx':")
print("multiplicar por 10 'c'")

# Creación da la operación
multi10_1 = (dataf1['c']*10)
print(multi10_1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C5'] = str(multi10_1)


# Ejercicio 3: head(n) - devuelve las n primeras filas del DataFrame
hojacalculo['B6'] = ("Primera fila del Dataframe del archivo 'dataframe1.xlsx':")
print("head")

# Creación da la operación
head1 = dataf1.head(1)
print(head1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C6'] = str(head1)


# Ejercicio 4: columns - devuelve una lista con los nombres de las columnas del DataFrame
hojacalculo['B7'] = ("Lista con los nombres de las columnas del DataFrame del archivo 'dataframe1.xlsx':")
print("columns")

# Creación da la operación
columns1 = dataf1.columns
print(columns1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C7'] = str(columns1)



# Guadaremos los ejercicios realizados
opera_doc_1.save('opera1_excel_1.xlsx')

print()

print("--------"*8)

print()

# Crearemos la variable para leer el excel 2, contiene números decimales y letras
num_2 = pd.read_excel("dataframe2.xlsx",sheet_name=0)
print(num_2,"\n")

# Esta variable recorre todo el Dataframe
dataf2 = pd.DataFrame(num_2)

# Procederemos a crear el archivo donde guardaremos las operaciones 

# Crear un archivo de Excel
opera_doc_2 = openpyxl.Workbook()

# Aquí asignamos una hoja de cálculo en blanco en el archivo creado en el paso anterior
hojacalculo2 = opera_doc_2.active

# Guardaremos lo que llevamos en el archivo con el nombre: 'opera2_excel_2.xlsx'
opera_doc_2.save('opera2_excel_2.xlsx')

# Ya creado, continuaremos diseñando algunas operaciones
hojacalculo2['A1'] = ("Operaciones realizadas con la librería pandas de Pyhton")
hojacalculo2['A3'] = ("Ejercicios parte 2")
hojacalculo2['A4'] = 1
hojacalculo2['A5'] = 2
hojacalculo2['A6'] = 3
hojacalculo2['A7'] = 4

# Ejercicio 1: mean() - devuelve la media de los datos del dataframe
hojacalculo2['B4'] = ("Media de los datos del Dataframe del archivo 'dataframe2.xlsx':")
print("mean")

# Creación da la operación
mean2 = dataf2.mean()
print(mean2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C4'] = str(mean2)


# Ejercicio 2: size - devuelve el número de elementos del DataFrame
hojacalculo2['B5'] = ("Número de elementos del DataFrame del archivo 'dataframe2.xlsx':")
print("size")

# Creación da la operación
size2 = dataf2.size
print(size2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C5'] = size2


# Ejercicio 3: dtypes - devuelve una serie con los tipos de datos de las columnas del DataFrame
hojacalculo2['B6'] = ("Tipos de datos de las columnas del DataFrame del archivo 'dataframe2.xlsx':")
print("dtypes")

# Creación da la operación
dtypes2 = dataf2.dtypes
print(dtypes2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C6'] = str(dtypes2)


# Ejercicio 4: tail(n) - devuelve las n últimas filas del DataFrame
hojacalculo2['B7'] = ("La última fila del Dataframe del archivo 'dataframe2.xlsx':")
print("tail")

# Creación da la operación
tail2 = dataf2.tail(1)
print(tail2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C7'] = str(tail2)



# Guadaremos los ejercicios realizados
opera_doc_2.save('opera2_excel_2.xlsx')
