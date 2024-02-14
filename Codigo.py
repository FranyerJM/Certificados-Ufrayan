#Desarrrollado por Franyer Marin 

#Programa certificados Ufrayan

#Instalar Librerias: pip install PyPDF2, pip install reportlab, pip install pandas, pip install PyMuPDF

from PyPDF2 import PdfWriter, PdfReader

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import io
import pandas as pd
import datetime

#Determinar Curso
print()
print("¿De que curso deseas realizar el certificado?")
print()
print("Si es un curso con pansantía, escribe 1")
print("Si es un curso corto, escribe 2")
print()
CursoPansatiaOCorto = int(input("Escribe el número: "))
print()
print("-------------------------------------")
print()
if CursoPansatiaOCorto == 1:
	print("¿Cual curso con pasantía?")
	print()
	print("Si es Farmacia, escribe 1")
	print("Si es Enfermería, escribe 2")
	print("Si es Bioanalisis, escribe 3")
	print("Si es Administración, escribe 4")
	print()
elif CursoPansatiaOCorto == 2:
	print("¿Cual curso corto?")
	print()
	print("Si es Computación, escribe 5")
	print("Si es Office, escribe 6")
	print("Si es Electronica, escribe 7")
	print("Si es Barbería, escribe 8")
	print("Si es Sistema de Uñas, escribe 9")
	print("Si es Depilación, escribe 10")
print()
Cursovar = int(input("Escribe el número: "))

#Instructor elejir usuario
print()

if CursoPansatiaOCorto == 2:  #Si es un Curso Corto solicitara aparecer el instructor o no
	print("-------------------------------------")
	print()
	print("¿Deseas que aparezca la sección instructor?")
	print()
	print("Si es SI, escribe 1")
	print("Si es NO, escribe 2")
	print()
	InstructorFirma = int(input("Escribe el número: "))

	
	if InstructorFirma == 1:  #si decide que si pedira si quieres que aparezca el nombre, si no simplemente el nombre saldra en blanco
		print()
		print("-------------------------------------")
		print()
		print("¿Deseas que aparezca el nombre del instructor?")
		print()
		print("Si es SI, escribe 1")
		print("Si es NO, escribe 2")
		print()
		InstruConfirName = int(input("Escribe el número: "))

		if InstruConfirName == 1:
			print()
			print("-------------------------------------")
			print()
			NombreInstructorVar = input("Escribe aquí el nombre del instructor: ")
		else:
			NombreInstructorVar = " " 
			print("El área del nombre en el instructor aparecera en blanco")
	else:
		InstruConfirName = 1
		InstruConfirName = 1
		NombreInstructorVar = " "
		print("No aparecera la sección del instructor")

else:
	InstructorFirma = 1
	InstruConfirName = 1
	NombreInstructorVar = " "

if Cursovar == 1:
	Curso = "Farmacia"
	TituloCurso = "Auxiliar de Farmacia"
elif Cursovar == 2:
	Curso = "Enfermeria"
	TituloCurso = "Auxiliar de Enfermería"
elif Cursovar == 3:
	Curso = "Bionalisis"
	TituloCurso = "Asistente de Laboratorio Clínico"
elif Cursovar == 4:
	Curso = "Administracion"
	TituloCurso = "Asistente Administrativo Contable"
elif Cursovar == 5:
	Curso = "Computacion"
	TituloCurso = Curso
elif Cursovar == 6:
	Curso = "Office"
	TituloCurso = Curso
elif Cursovar == 7:
	Curso = "Electronica"
	TituloCurso = Curso
elif Cursovar == 8:
	Curso = "Barberia"
	TituloCurso = "Barbería"
elif Cursovar == 9:
	Curso = "Sistema de Uñas"
	TituloCurso = Curso
elif Cursovar == 10:
	Curso = "Depilacion"
	TituloCurso = "Depilación Facial"


CursoMayus = Curso.upper() #Convierte el curso en mayusculas

#Leer datos
Datos = pd.read_excel('Datos/datos.xlsx')

for i in range(len(Datos)):

	# Crear variables
	Estudiante = Datos.loc[i, 'Nombre']
	Cedula = Datos.loc[i, 'Cedula']

	#Fecha del finalizacion del curso
	FechaHoy = datetime.datetime.now() #Sacar Fecha
	FechaModi = FechaHoy.strftime("%m,%Y") 

	seccion = FechaModi.split(",")
	MesNum = int(seccion[0])
	Año = seccion[1]

	MesesEspañol = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

	NombreMes = MesesEspañol[MesNum-1]
	MesMinu = NombreMes.lower()
	Estado = "Edo. Carabobo, "
	FechaTotal = MesMinu + " " + Año
	FinaCur = Estado + FechaTotal
	MesMayu = NombreMes.upper()
	FechaTotalMayus = MesMayu + " " + Año

	#Nombre del instructor si existe
	#Existira si es un curso con pansantia y si el usuario lo decide en cursos cortos
	#Para los cursos con pansantias
	if Cursovar == 1:
		Instructor = "Dr. Lee Nuñez"
	elif Cursovar == 2:
		Instructor = "Lic. Lisbeth De Logan"
	elif Cursovar == 3:
		Instructor = "Lic. Luis Marcano"
	elif Cursovar == 4:
		Instructor = ""

	else:
		InstruConfirName == 1
		Instructor = NombreInstructorVar

	#Duracion del curso
	if Cursovar == 1 or 2 or 3 or 4:
		Duracion = "280"
	if Cursovar == 5 or 6 or 7 or 8 or 9 or 10:
		Duracion = "36"


	# Crear marca de agua
	packet = io.BytesIO()
	width, height = letter #Esta predefinido para leer archivos de 612 x 792
	c = canvas.Canvas(packet, pagesize=(width*5.25, height*3.14)) #multiplico: Ancho: 612*5.25 = 3213px __ Alto: 792*3.14 = 2289px


	# Obtener tipo de letra
	reportlab.rl_config.warnOnMissingFontGlyphs = 0
	pdfmetrics.registerFont(TTFont('Edward', 'font/Edwardian.ttf')) # Para el nombre del estudiante
	pdfmetrics.registerFont(TTFont('MyriadBold', 'font/MYRIADPRO-BOLD.ttf'))
	pdfmetrics.registerFont(TTFont('Myriad', 'font/MYRIADPRO-REGULAR.ttf'))
	pdfmetrics.registerFont(TTFont('TimeNewBold', 'font/times new roman bold.ttf')) # Para la cedula


	# Ubico las variables en el espacio

	# Nombre - Edward
	c.setFillColorRGB(0,0,0)
	c.setFont('Edward', 72)
	c.drawCentredString(410, 300, Estudiante)

	# Cedula - Myriad
	c.setFillColorRGB(0,0,0)
	c.setFont('MyriadBold', 14.5)
	c.drawCentredString(375, 230, Cedula)

	# Finalizacion del curso (Pie de Pagina)
	c.setFillColorRGB(0,0,0)
	c.setFont('TimeNewBold', 16)
	c.drawCentredString(390, 35, FinaCur)

	# Instructor
	c.setFillColorRGB(0,0,0)
	c.setFont('Myriad', 14.1)
	c.drawCentredString(390, 87, Instructor)

	# Duracion
	c.setFillColorRGB(0,0,0)
	c.setFont('Myriad', 12.1)
	c.drawCentredString(673, 163, Duracion)

	# Titulo del curso (Nombre del curso)
	c.setFillColorRGB(0,0,0,)
	c.setFont('MyriadBold', 17.5)
	c.drawCentredString(410, 211, TituloCurso)

	c.save()


	# Invoco plantilla PDF

	if (InstructorFirma == 1) or (CursoPansatiaOCorto == 1):
		existing_pdf = PdfReader(open("Plantillas/Plantilla certificados General.pdf", "rb"))
	elif InstructorFirma == 2:
		existing_pdf = PdfReader(open("Plantillas/Plantilla certificados General sin instructor.pdf", "rb"))
	
	page = existing_pdf.pages[0]


	# Añadir Marca de agua al PDF
	packet.seek(0)
	new_pdf = PdfReader(packet)
	page.merge_page(new_pdf.pages[0])


	#Exportar Certificado
	file_name = Estudiante

	# Personalizar cada nombre del archivo
	certificado = "Certificados/Certificados " + Curso +"/" + file_name + " " + Curso + ".pdf"

	outputStream = open(certificado, "wb")
	output = PdfWriter()
	output.add_page(page)
	output.write(outputStream)
	outputStream.close()

# Unir PDf
import os
import fitz

#seccion que hace el trabajo de unir

def unir_pdf(carpeta_entrada, archivo_salida):
	pdf_salida = fitz.open() #crea el archivo

	#leer PDf'S
	for nombre_archivo in os.listdir(carpeta_entrada):
		if nombre_archivo.endswith('.pdf'): #Seleccionar solo archivos PDF
			ruta_archivo = os.path.join(carpeta_entrada, nombre_archivo) #ubico

			pdf = fitz.open(ruta_archivo) #abrir archivo

			pdf_salida.insert_pdf(pdf)

			pdf.close()

	pdf_salida.save(archivo_salida)
	pdf_salida.close()

# Eliminar Pdf's

def eliminar_pdf(carpeta_elim, pdfs_elim):
	for archivo in pdfs_elim:
		rutaTotal = os.path.join(carpeta_elim, archivo)
		if archivo.endswith(Curso+'.pdf'): #solo los pdf
			if os.path.isfile(rutaTotal):
				os.remove(rutaTotal)

print()
print("-----------------------------------------------------")
print()
print("Los certificados se han creado correctamente")
print()
print("Se ha creado un PDF por cada certificado")
print()
print("¿Deseas unir los PDF's?")
print()
print("Si es SI, escribe 1:")
print("Si es NO, escribe 2:")
print()
PDFVar = int(input("Escribe aquí: "))

#Determinar si se unen o no
if PDFVar == 1:
	carpeta_entrada = "Certificados/Certificados " + Curso + "/"

	archivo_salida = "Certificados/Certificados " + Curso + "/" + "Certificados " + Curso + " Comprimido/"+ CursoMayus + " " + FechaTotalMayus + ".pdf"

	#Uno los PDF
	unir_pdf(carpeta_entrada, archivo_salida)
	print("--------------------------------------------------------------------")
	print()
	print("Listo, tu archivo PDF está en la carpeta: Certificados",Curso,"/","Certificados",Curso,"Comprimido")
	print()
	print("Quedaron los PDF's individuales, ¿deseas eliminarlos?") #interactuar a ver si quiere eliminar los PDFs
	print()
	print("Si es SI, escribe 1")
	print("Si es NO, escribe 2")
	EliminarPDFSIoNO = int(input("Escribe aquí: "))

	if EliminarPDFSIoNO == 1:
		carpeta_elim = "Certificados/Certificados " + Curso + "/"
		pdfs_elim = os.listdir(carpeta_elim)
		eliminar_pdf(carpeta_elim, pdfs_elim)
		print()
		print("Los PDF's individuales se eliminaron correctamente")
	elif EliminarPDFSIoNO == 2:
		print()
		print("Los PDF's individuales se conservaran en la carpeta: "+"Certificados/Certificados " + Curso)

else:
	print("Listo, no se unieron y tus PDF's estan gurdadados en la carpeta: Certificados/Certificados",Curso)
print()
input("Presiona enter para finalizar: ")