import tkinter as tk
from tkinter import filedialog
import docx
import os
import comtypes.client
import shutil
from docx import Document


def merge_docx_files(origen, destino):
    doc_origen = docx.Document(origen)
    doc_destino = docx.Document(destino)

    for element in doc_origen.element.body:
        doc_destino.element.body.append(element)
    doc_destino.save(destino)
    os.remove(origen)

def get_docx_files(path):
    print("get_docx_files")
    docx_files = []
    for file in os.listdir(path):
        if file.endswith(".docx") and file != 'destination.docx' and "2023" not in file:
            docx_files.append(os.path.join(path, file))
    return docx_files

def get_doc_files(path):
    print("get_docx_files")
    docx_files = []
    for file in os.listdir(path):
        if file.endswith(".doc") and file != 'destination.docx' and "2023" not in file:
            docx_files.append(os.path.join(path, file))
    return docx_files

def get_folder_paths(path):
    print("get_folder_paths")
    #Entrega una lista de todos los paths hacia todas las posibles carpetas dentro de una carpeta específica.
    
    folder_paths = [folder_path2]
    for root, dirs, files in os.walk(path):
        for dir in dirs:
            folder_paths.append(os.path.join(root, dir))
    return folder_paths

def doc_to_docx(source_file, destination_file):
    """
    Transforma un archivo .doc en un .docx
    """
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(source_file)
    doc.SaveAs(destination_file, FileFormat=16)
    doc.Close()
    word.Quit()
    os.remove(source_file)

def procesar(folder_path2):
    btn_process.config(bg="red", text="Procesando", state="disable")
    contador = 0
    print(folder_path2)
    
    docx_path = get_folder_paths(folder_path2)
    print(docx_path)

    print("empieza a ejecutarse")
    for path2 in docx_path:
        print(path2)
    
        docx_files = get_doc_files(path2)
    
        for name in docx_files:
            print(name)
            
            path3 =f"{name}".replace(".doc", ".docx")
            doc_to_docx(name, path3)
            print(f"{name} como .docx")
    

    docx_path = get_folder_paths(folder_path2)
    print(docx_path)
    for path2 in docx_path:
        docx_files = get_docx_files(path2)
        print(docx_files)
        for name in docx_files:
            print(name)
            new_name = f"{name}".replace(".docx", " 2023.docx")
            shutil.copy2('D:\\ATestCode\\ProgramaDocxIunge\\destination.docx', new_name)

            source_file = name
            
            merge_docx_files(source_file, new_name)
            contador = contador + 1

    print(f"{contador} archivos procesados")




def select_folder():
    global folder_path2
    # Abrimos una ventana de selección de archivos
    folder_path2 = filedialog.askdirectory()
    folder_path2 = folder_path2.replace("/", "\\")
    # Guardamos la ruta en una variable
    selected_folder.set(folder_path2)
    # Mostramos el path en la consola
    print(folder_path2)
    
    
    

# Creamos una ventana principal
root = tk.Tk()
root.geometry("400x300+0+0")
root.title("Seleccionar carpeta")

# Creamos una etiqueta para mostrar el path seleccionado
selected_folder = tk.StringVar()
label = tk.Label(root, textvariable=selected_folder)
label.pack()

# Creamos un botón para abrir la ventana de selección
button = tk.Button(root, text="Seleccionar carpeta", command=select_folder)
button.pack()
btn_process = tk.Button(root, text="Procesar", bg="green", command=lambda: procesar(folder_path2))
btn_process.pack()
root.mainloop()
