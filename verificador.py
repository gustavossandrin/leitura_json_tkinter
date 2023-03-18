import os
from tkinter import *
from tkinter import Tk, messagebox
import json
import openpyxl
import smtplib
from email.message import EmailMessage


co0 = "#f0f3f5"  # Preta / black
co1 = "#feffff"  # branca / white
co2 = "#3fb5a3"  # verde / green
co3 = "#38576b"  # valor / value
co4 = "#403d3d"  # letra / letters


# messagebox.showwarning('Error', 'Arquivo não encontrado')



def encontrar_falhas():
    arquivo_nome = e_arquivo.get()
    email = e_email.get()

    enviou_email = False

    if os.path.isfile(arquivo_nome):
        book = openpyxl.Workbook()
        initial_page = book['Sheet']
        initial_page.append(['Mensagens que falharam'])
        initial_page.append(['ID', 'Dia', 'Hora'])

        with open(arquivo_nome, encoding='utf-8') as arquivo:
            dados = json.load(arquivo)
            mensagens = dados['messages']
            for msg in mensagens:
                try:
                    falhou = msg['text']
                    if falhou.endswith('FALHOU'):
                        id_msg = msg['id']
                        dia_e_hora = msg['date']
                        dia_e_hora = dia_e_hora.split('T')
                        dia_msg = dia_e_hora[0]
                        hora_msg = dia_e_hora[1]
                        initial_page.append([id_msg, dia_msg, hora_msg])
                except:
                    pass
            book.save('mensagem_que_falharam.xlsx')

            if os.path.isfile('mensagem_que_falharam.xlsx') and email:
                email_msg = EmailMessage()
                email_msg['Subject'] = 'Relatório de falhas'
                email_msg['From'] = 'gustavosandrindattein123@gmail.com'
                email_msg['To'] = email
                email_msg.set_content("Segue o relatorio")

                with open('mensagem_que_falharam.xlsx', 'rb') as content_file:
                    content = content_file.read()
                    email_msg.add_attachment(content, maintype='application', subtype='xlsx', filename='mensagem_que_falharam.xlsx')

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login('gustavosandrindattein123@gmail.com', 'senha email')
                    smtp.send_message(email_msg)
                    enviou_email = True
                    messagebox.showinfo('Success', 'Email Enviado com sucesso!')

            if not enviou_email:
                messagebox.showinfo('Success', 'Relatório Criado!')

            janela.destroy()



    else:
        messagebox.showwarning('Error', 'Arquivo não encontrado')


# Criando Janela
janela = Tk()
janela.title('')
janela.geometry('310x300')
janela.configure(background=co1)
janela.resizable(width=FALSE, height=FALSE)

# Dividindo a Janela
frame_cima = Frame(janela, width=310, height=50, bg=co1, relief='flat')
frame_cima.grid(row=0, column=0, pady=1, padx=0, sticky=NSEW)

frame_baixo = Frame(janela, width=310, height=250, bg=co1, relief='flat')
frame_baixo.grid(row=1, column=0, pady=1, padx=0, sticky=NSEW)

# Configurando frame cima ----------------
l_nome = Label(frame_cima, text="VERIFICAR FALHAS", anchor=NE, font=('Ivy 18'), bg=co1, fg=co4)
l_nome.place(x=5, y=10)

l_linha = Label(frame_cima, text="", width=235, anchor=NW, font=('Ivy 1'), bg=co2, fg=co4)
l_linha.place(x=10, y=45)

# Configurando frame baixo ----------------
l_arquivo = Label(frame_baixo, text="Nome arquivo*", anchor=NW, font=('Ivy 10'), bg=co1, fg=co4)
l_arquivo.place(x=10, y=20)
e_arquivo = Entry(frame_baixo, width=25, justify='left', font=("", 15), highlightthickness=1, relief='solid')
e_arquivo.place(x=14, y=50)

l_email = Label(frame_baixo, text="Email", anchor=NW, font=('Ivy 10'), bg=co1, fg=co4)
l_email.place(x=10, y=95)
e_email = Entry(frame_baixo, width=25, justify='left', font=("", 15), highlightthickness=1, relief='solid')
e_email.place(x=14, y=130)

b_verificar = Button(frame_baixo, text="Verificar", width=39, height=2, font=('Ivy 8 bold'), bg=co2, fg=co1,
                     relief=RAISED, overrelief=RIDGE, command=encontrar_falhas)
b_verificar.place(x=15, y=180)

janela.mainloop()
