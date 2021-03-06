from tkinter import *
import csv
import xlrd

#dividir em firstname e lastname
def divideNome(nome):
        cont = 0
        firstname = ''
        lastname = ''
        for i in range(len(nome)):
            if nome[i] != ' ' and cont == 0:
                firstname+=nome[i]
                
            if cont >= 1:
                lastname+=nome[i]
                
            if nome[i] == ' ':
                if cont == 0:
                    firstname+=';'
                cont+=1
                
            if i == (len(nome)-1):
                lastname+=';'
                
        lastname = lastname.title()
        lastname = lastname.split('  ')
        firstname = firstname.capitalize()
        
        return firstname,lastname

#escrever dados
def escrever(arquivo,dados):
        with open(arquivo+'.csv', 'w', newline='') as csvfile:
            spamwriter = csv.writer(csvfile, delimiter=' ', quotechar=' ')
            string = "username;"+"password;"+"firstname;"+"lastname;"+"email;"+"course1;"+"group1;"
            
            lista1 = []
            lista1.append(string)
            
            lista = []
            lista.append(dados)

            spamwriter.writerow(lista1)
            spamwriter.writerow(lista)

#ler dados
def ler(arquivo_abrir,arquivo_dest,curso):
        workbook = xlrd.open_workbook(arquivo_abrir+'.xlsx')
        sheet = workbook.sheet_by_index(0)
        parametros = []
        dados = []
        string = ""
        cont = 1
        grupo = ';'
        
        for i in range(sheet.ncols): #numero de colunas
            parametros.append(sheet.cell_value(0, i))

        while(cont != sheet.nrows):
            for i in range(len(parametros)):
                if parametros[i] == 'cpf' or parametros[i] == 'CPF':
                    username = str(int(sheet.cell_value(cont, i))) +';'
                
                if parametros[i] == 'nome' or parametros[i] == 'estudante' or parametros[i] == 'Nome do Aluno' or parametros[i] == 'Nome Aluno' or parametros[i] == 'Nome completo':
                        nome = divideNome(sheet.cell_value(cont, i))
                        firstname = nome[0]
                        lastname = nome[1][0]
                        print(lastname)
               
                if parametros[i] == 'email' or parametros[i] == 'E-mail' or parametros[i] == 'Email Aluno' or parametros[i] == 'e-mail':
                    email = sheet.cell_value(cont, i) +';'
                    email = email.lower()

                if parametros[i] == 'Código Turma':
                        grupo = sheet.cell_value(cont, i)+';'
                    
                password = 'Senha;'
                
                if i == (len(parametros)-1):
                    string += username + password + firstname + lastname + email + curso+ ';' + grupo + ';\n'
                    escrever(arquivo_dest,string)
                
            if sheet.nrows > 2: #se linha maior que 2
                 cont +=1
            else:
                break

class Janela:
    def __init__(self,top):
        self.frame0 = Frame(top)
        self.frame0.pack()

        self.frame1 = Frame(top)
        self.frame1.pack()

        self.frame2 = Frame(top)
        self.frame2.pack()
        
        self.frame3 = Frame(top)
        self.frame3.pack()

        self.entradaLabel = Label(self.frame0,text='Nome do arquivo de entrada: ',font=('Arial',12))
        self.entradaLabel.pack(side=LEFT)
        
        self.entrada = Entry(self.frame0,width = 20,font = ('Arial',17))
        self.entrada.pack(side=RIGHT)

        self.espaco = Label(self.frame0,text='')
        self.espaco['height'] = 4
        self.espaco.pack()

        self.saidaLabel = Label(self.frame1,text='Nome do arquivo de saida:     ',font=('Arial',12))
        self.saidaLabel.pack(side=LEFT)
        
        self.saida = Entry(self.frame1,width = 20,font = ('Arial',17))
        self.saida.pack(side=RIGHT)

        self.espaco = Label(self.frame1,text='')
        self.espaco['height'] = 4
        self.espaco.pack()

        self.cursoLabel = Label(self.frame2,text='Curso:                                        ',font=('Arial',12))
        self.cursoLabel.pack(side=LEFT)
        
        self.curso = Entry(self.frame2,width = 20,font = ('Arial',17))
        self.curso.pack(side=RIGHT)

        self.espaco = Label(self.frame2,text='')
        self.espaco['height'] = 3
        self.espaco.pack()

        self.finalizar = Button(self.frame3,text='Finalizar',font = ('Arial',12),command=self.finalizar,height = 1,width = 9)
        self.finalizar.bind("<Return>",self.finalizar)
        self.finalizar.pack()
        
    def finalizar(self):
        arquivo_abrir = self.entrada.get()
        arquivo_dest = self.saida.get()
        curso = self.curso.get()
        
        ler(arquivo_abrir,arquivo_dest,curso)

        
raiz = Tk()
raiz.title('CEAD - CSV 1.0')
Janela(raiz)
raiz.mainloop()
