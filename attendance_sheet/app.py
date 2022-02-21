from tkinter import Tk, StringVar, Label, Button, Entry, filedialog, W
from os.path import exists
import generator as gp

def cmdExec():
    if checkFileExist(textIn.get()) and checkFileExist(textOut.get()):
        result.set("Gerando planilha de presença ...")
        isSuccess = gp.main(textIn.get(), textOut.get())
        if isSuccess:
            result.set("Planilha de presença gerada com successo")
        else:
            result.set("Ocorreu um erro ao tentar gerar a planilha de presença")

def checkFileExist(file):
    if(file != ""):
        name  = file.split('/')
        if exists(file):
            return True
        result.set(f"Erro: Arquivo de {name[-1]} não encontrado")
        return False
    result.set("Erro: Campo vazio")
    return False
    

def cmdSearchFileIn():
    filename = filedialog.askopenfilename()
    result.set("")
    textIn.set(filename)

def cmdSearchFileOut():
    filename = filedialog.askopenfilename()
    result.set("")
    textOut.set(filename)
    
screen = Tk()
screen.title("Gerador de Lista de Presença")
textIn = StringVar()
textOut = StringVar()
result = StringVar()

# pos screen
width, height = 500, 200
widthScreen = screen.winfo_screenwidth()
heightScreen = screen.winfo_screenheight()

posX = int(widthScreen/2 - width/2)
posY = int(heightScreen/2 - height/2)

screen.geometry(f"{width}x{height}+{posX}+{posY}")

# Labels
labelFileIn = Label(screen, text="Escolha o arquivo de entrada:").grid(row=0, sticky=W)
labelFileOut = Label(screen, text="Escolha o arquivo de saída:").grid(row=2, sticky=W)
labelResult = Label(screen, textvariable=result).grid(row=5, pady=10)

# Text box
textBoxFileIn = Entry(screen, textvariable=textIn).grid(row=1, padx=5, pady=5 ,ipadx=120)
textBoxFileOut = Entry(screen, textvariable=textOut).grid(row=3, padx=5, pady=5, ipadx=120)
# Butões
btnSearchFileIn = Button(screen, text= "Buscar", command=cmdSearchFileIn).grid(row=1, column=1)
btnSearchFileOut = Button(screen, text= "Buscar", command=cmdSearchFileOut).grid(row=3, column=1)
btnExec = Button(screen, text= "Executar", command=cmdExec).grid(row=4, pady=10)

screen.mainloop()