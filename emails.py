# -*- coding: utf-8 -*-
import os
import pickle
import re

import gspread as gs
import unidecode as unidecode
from google.auth.transport import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

ARCHIVO = "Prueba de Emails"
INICIO_ARCHIVO = 1

DOMINIO = "@colegio-altamira.cl"


def celda_is_null(cell):
    if cell == "" or cell is None:
        return True
    return False


def ask_default_pass():
    while True:
        pass1 = input("\nIndique la contraseña de las cuentas nuevas: ")
        pass2 = input("Confirme contraseña: ")

        if pass1 == pass2:
            return pass1

        print("Las contraseñas no coinciden")


def generate_adresses(names_o):
    names = names_o.copy()
    emails_creados = []

    for curso in names.values():
        for estudiantes in curso:
            x = 3
            while True:
                email = unidecode.unidecode(estudiantes[0].strip(" ") + "." + estudiantes[1][:x]).lower() + DOMINIO
                if email not in emails_creados:
                    estudiantes.append(email)
                    emails_creados.append(email)
                    break

                x += 1
                if x >= len(estudiantes[1]):
                    print("\nERROR: Imposible resolver el email para %s %s" % (estudiantes[0], estudiantes[1]))
                    exit(1)
    return names


print("---------------------------------------")
print("      Generador de Correos v1.0.0")
print("---------------------------------------")
print()

print("INFO: Iniciando..")

# Conseguimos las credenciales para usar la API de Google Drive
SCOPE = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/admin.directory.user']
SCOPE_GSUITE = ['https://www.googleapis.com/auth/admin.directory.user']

print("INFO: Estableciendo conexión con los servidores..")
print("INFO: Autenticando..")
creds = ServiceAccountCredentials.from_json_keyfile_name('creds_sheets.json', SCOPE)
cliente = gs.authorize(creds)

print("INFO: Consiguiendo la lista de correos para generar..")
lista_estudiantes = None
try:
    archivo = cliente.open(ARCHIVO).sheet1
    lista_estudiantes = archivo.col_values(INICIO_ARCHIVO)
except gs.exceptions.SpreadsheetNotFound:
    print("\nERROR: No se ha encontrado el archivo")
    exit(1)

print("INFO: Procesando datos..")
curso = lista_estudiantes[0]
estudiantes_por_curso = {curso: []}
print("\nCurso => %s:" % curso)

cursor = 0
try:
    while True:
        cursor += 1
        if len(lista_estudiantes) <= cursor:
            break
        celda = lista_estudiantes[cursor]

        if celda_is_null(celda):
            if len(lista_estudiantes) <= cursor + 1:
                break

            prox_celda = lista_estudiantes[cursor + 1]
            if celda_is_null(prox_celda):
                break

            re_curso = "((?:[1-7]{1}|I{1,3}|IV)º[A-D]{1})"
            match_curso = re.search(re_curso, prox_celda)

            if not match_curso:
                print("\nERROR: Caso no esperado en el archivo")
                print("ERROR: Se esperaba curso o nada, pero se encontró \"%s\"" % prox_celda)
                exit(1)

            cursor += 1
            curso = match_curso.string
            estudiantes_por_curso[curso] = []
            print("\nCurso => %s:" % curso)
            continue

        apellido_estudiante, nombre_estudiante = celda.rstrip().split(", ")
        print("\t%s %s" % (nombre_estudiante, apellido_estudiante))

        estudiantes_por_curso[curso].append([nombre_estudiante, apellido_estudiante])

except TypeError:
    print("\nERROR: Excepción al intentar descomprimir la celda")
    exit(1)

print("\nINFO: Iniciando conexión con Google Admin...")
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'creds_gsuite.json', SCOPE_GSUITE)
        creds = flow.run_local_server(port=0)

    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('admin', 'directory_v1', credentials=creds)
print("INFO: Conexión establecida")

default_pass = ask_default_pass()

print("INFO: Generando correos..")

for curso, estudiantes in generate_adresses(estudiantes_por_curso).items():
    for estudiante in estudiantes:
        print("\n\tCreando cuenta para %s %s, correo: %s" % (estudiante[0], estudiante[1], estudiante[2]))
        user = {
            "name": {
                "familyName": estudiante[1],
                "givenName": estudiante[0],
            },
            "password": default_pass,
            "changePasswordAtNextLogin": True,
            "primaryEmail": estudiante[2],
        }

        result = service.users().insert(body=user).execute()

        print("\tCuenta creada con el ID: %d" % result["id"])

print("\nINFO: Listo!")
print("\nINFO: Terminando")
exit(0)
