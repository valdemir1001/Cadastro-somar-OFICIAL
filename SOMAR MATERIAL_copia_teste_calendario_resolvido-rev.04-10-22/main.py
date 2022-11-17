import os

from IPython.terminal.pt_inputhooks import tk

from modulos import *
from funcionalidades import *
from openpyxl.workbook import Workbook
from openpyxl import *
from Listas_geral import *
import base64
import customtkinter


root = Tk()

class Application(Funcionalidades,Validadores):
# Inicio
    def __init__(self,master=None):
        self.root = root
        self.imagens_base64()

        self.python_image = True
        self.validaEntrada()
        self.tela()
        self.frames_da_tela()
        self.imag()
        self.widgets_frame1()
        self.lista_frame2()
        self.montaTabelas()
        self.select_lista()
        self.menus()
        root.mainloop()

# Itens da Tela
    def tela(self):
        img = PhotoImage(data=self.ico)
        self.root.tk.call('wm', 'iconphoto', root._w, img)
        self.root.title('CADASTRO DE MATERIAL SOMAR')
        self.root.configure(background='#001f36')
        self.root.geometry('1200x1000+10+10')
        self.root.resizable(True, True)
        self.root.maxsize(width=1000, height=900)
        self.root.minsize(width=700, height=600)
# Variáveis dos Widgets
        self.root.font = ('Verdana',9,'bold')
        self.root.font_cor = 'black'
        self.root.bg_red = '#e72313'
        self.root.color_aba1 = '#1e8c93'
        self.root.color_aba2 = '#079ea6'
        self.root.color_aba3 = '#00b9bd'
        self.root.color_aba4 = '#9e0c39'
        self.root.cor_inter_aba3 = '#001f36'
        self.root.cor_bord_aba3 = '#004853'
# Frames Principais que fica dentro das Abas
    def frames_da_tela(self):
    # Frame da Logo da Imagem da Somar
        self.frame_3 = Frame(self.root)
        self.frame_3.place(relx=0.01, rely=0.01, relwidth=0.08, relheight=0.05)
    # Frame 1 da aba 1
        self.frame_1 = Frame(self.root, border=4, bg='#96d7eb', highlightbackground='#1c5560', highlightthickness=3)
        self.frame_1.place(relx=0.0, rely=0.07, relwidth=1, relheight=0.48)
    # Frame 2 da aba1
        self.frame_2 = Frame(self.root, border=4, bg='#96d7eb', highlightbackground='#1c5560', highlightthickness=3)
        self.frame_2.place(relx=0.0, rely=0.55, relwidth=1, relheight=0.44)

    # __________________________________________Cabeçalho da Tela___________________________

        self.lb_principal = Label(self.root, text='Controle da Entrada e Saída de Material', font=('Verdana', 20, 'bold'),
                                  bg='#001f36', fg='white')
        self.lb_principal.place(relx=0.10, rely=0.01, relwidth=0.89, relheight=0.05)
# Widgets com suas Canvas Botões e Label
    def widgets_frame1(self):
# Criando ABAS
        self.abas = ttk.Notebook(self.frame_1)
        self.aba1 = Frame(self.abas)
        self.aba2 = Frame(self.abas)
# ABAS
        self.aba1.configure(background='#96d7eb')
        self.aba2.configure(background='#96d7eb')
# Nomeia as Abas
        self.abas.add(self.aba1,text='CONTROLE DE SAÍDA DE MATERIAL')
        self.abas.add(self.aba2,text='CADASTRO DE MATERIAL E ENDEREÇOS')
# Posiciona as ABAS
        self.abas.place(relx=0,rely=0,relwidth=1,relheight=1)

# CABEÇALHO CONTROLE DE SAÍDA DE MATERIAL
    # CANVAS CABEÇALHO MATERAL
        self.canvas_cab_mat = Canvas(self.aba1, bd=0,bg=self.root.cor_inter_aba3, highlightbackground=self.root.cor_bord_aba3,highlightthickness=5)
        self.canvas_cab_mat.place(relx=0.02, rely=0.02, relwidth=0.97, relheight=0.12)

        self.lb_cab_mat = Label(self.aba1,text='Cadastro de Material',font=('Verdana', 15, 'bold'),bg='#001f36', fg='white')
        self.lb_cab_mat.place(relx=0.38, rely=0.05, relwidth=0.25, relheight=0.06)

# Botão CADASTRAR
        self.bt_novo = customtkinter.CTkButton(self.aba1, command = self.add_valores,
                                                        text = 'cadastrar'.upper(), text_font = ('Arial 7 bold'),text_color = "white",
                                                        width = 900, compound = BOTTOM, relief = RAISED, 
                                                        hover = True, hover_color = "blue",
                                                        border_width = 2, border_color = "#0e002f", corner_radius = 8,
                                                        fg_color= "#092b5a"
                                                )
        self.bt_novo.place(relx=0.22, rely=0.18, relwidth=0.09, relheight=0.10)

# Botão ALTERAR
        self.bt_alterar = customtkinter.CTkButton(self.aba1, command = self.alter_valores,
                                                        text = 'Alterar'.upper(), text_font = ('Arial 7 bold'),text_color = "white",
                                                        width = 900, compound = BOTTOM, relief = RAISED, 
                                                        hover = True, hover_color = "blue",
                                                        border_width = 2, border_color = "#0e002f", corner_radius = 8,
                                                        fg_color = "#092b5a"
                                                )
        self.bt_alterar.place(relx=0.32, rely=0.18, relwidth=0.09, relheight=0.10)
# Botao LIMPAR
        self.bt_limpar = customtkinter.CTkButton(self.aba1, command = self.limpa_tela,
                                                        text = 'Limpar'.upper(), text_font = ('Arial 7 bold'), text_color = "white",
                                                        width = 900, compound = BOTTOM, relief = RAISED, 
                                                        hover = True, hover_color = "blue",
                                                        border_width = 2, border_color = "#0e002f", corner_radius = 8,
                                                        fg_color= "#092b5a",
                                                )
        self.bt_limpar.place(relx=0.42, rely=0.18, relwidth=0.09, relheight=0.10)

# Botão Excluir
        self.bt_apagar =customtkinter.CTkButton(self.aba1, command = self.deleta_valores,
                                                        text = 'Excluir'.upper(),text_font = ('Arial 7 bold'),text_color = "white",
                                                        width = 900, compound = BOTTOM, relief = RAISED, 
                                                        hover = True, hover_color = "blue",
                                                        border_width = 2,border_color = "#0e002f", corner_radius = 8,
                                                        fg_color = "#092b5a",
                                                )
        self.bt_apagar.place(relx=0.52, rely=0.18, relwidth=0.09, relheight=0.10)
# ID_____________

# Canvas do ID
        self.canvas_bt = customtkinter.CTkFrame(self.aba1,  
                                                        width = 200, height = 200,
                                                        border_width = 2, border_color = "#0e002f", corner_radius = 10,
                                                        fg_color='#2e97b7'
                                                )
        self.canvas_bt.place(relx=0.02,rely=0.15,relwidth=0.10,relheight=0.15)
        
# LABEL E ENTRY do ID
        self.lb_id = Label(self.aba1, text='ID', font=(self.root.font,11,'bold'), bg='#2e97b7', fg=self.root.font_cor)
        self.lb_id.place(relx=0.03, rely=0.17,relheight=0.04)
        
        self.id_entry = customtkinter.CTkEntry(self.aba1, validate = 'key', validatecommand = self.vcmd2,
                                                        text_font = self.root.font,
                                                        width = 120, height = 25,
                                                        border_width = 2, border_color = 'black', corner_radius = 8,
                                                        bg_color = '#2e97b7'
                                                )
        self.id_entry.place(relx=0.03, rely=0.22, relwidth=0.08, height=21)

# Canvas das Labels e Entrys
        self.canvas_bt = customtkinter.CTkFrame(self.aba1,        
                                                        width = 200, height = 200,
                                                        border_width = 2, border_color = "#0e002f", corner_radius = 10,
                                                        fg_color='#2e97b7'
                                                )
        self.canvas_bt.place(relx=0.02,rely=0.31,relwidth=0.59,relheight=0.66)
        
# Label e Entry DATA
        self.lb_data = Label(self.aba1, text='DATA', font=self.root.font, bg='#2e97b7', fg=self.root.font_cor)
        self.lb_data.place(relx=0.06, rely=0.35,relwidth=0.05, relheight=0.05)
        self.data_entry = Entry(self.aba1,font=self.root.font)
        self.data_entry.place(relx=0.12, rely=0.35, relwidth=0.12, relheight=0.06)
# Label e Entry Endereço
        self.lb_endereco = Label(self.aba1, text='ENDEREÇO', font=self.root.font, bg='#2e97b7', fg=self.root.font_cor)
        self.lb_endereco.place(relx=0.03, rely=0.45,relwidth=0.08, relheight=0.05)
        self.endereco_entry = Entry(self.aba1,font=self.root.font)
        self.endereco_entry.place(relx=0.12, rely=0.45, relwidth=0.30, relheight=0.06)
# Label e Entry Material
        self.lb_material = Label(self.aba1, text='MATERIAL',font=self.root.font, bg='#2e97b7',fg=self.root.font_cor)
        self.lb_material.place(relx=0.03, rely=0.55,relwidth=0.08, relheight=0.05)
        self.material_entry = ttk.Combobox(self.aba1,font=self.root.font,values=lista_material)
        self.material_entry.place(relx=0.12, rely=0.55, relwidth=0.30, relheight=0.06)
# Label e Entry QUANTIDADE
        self.lb_quantidade = Label(self.aba1, text='QUANT.',font=self.root.font, bg='#2e97b7',fg=self.root.font_cor)
        self.lb_quantidade.place(relx=0.05, rely=0.65,relwidth=0.06, relheight=0.05)
        self.quantidade_entry = ttk.Combobox(self.aba1,font=self.root.font,values=lista_Quant)
        self.quantidade_entry.place(relx=0.12, rely=0.65, relwidth=0.08, relheight=0.06)
# Label e Entry Unidade de Medida
        self.lb_uni = Label(self.aba1, text='UNI. MEDIDA.',font=self.root.font, bg='#2e97b7',fg=self.root.font_cor)
        self.lb_uni.place(relx=0.22, rely=0.65,relwidth=0.10, relheight=0.05)
        self.uni_entry = ttk.Combobox(self.aba1,font=self.root.font,values=lista_uni)
        self.uni_entry.place(relx=0.33, rely=0.65, relwidth=0.09, relheight=0.06)
# Label e Entry OS
        self.lb_os = Label(self.aba1, text='O.S.', font=('Verdana',12,'bold'), bg='#2e97b7', fg=self.root.font_cor)
        self.lb_os.place(relx=0.06, rely=0.75,relwidth=0.05, relheight=0.05)
        self.os_entry = Entry(self.aba1,font=self.root.font)
        self.os_entry.place(relx=0.12, rely=0.75, relwidth=0.10, relheight=0.06)
# Label e Entry Encarregado
        self.lb_encarregado = Label(self.aba1, text='ENCARREGADO',font=self.root.font, bg='#2e97b7', fg=self.root.font_cor)
        self.lb_encarregado.place(relx=0.03, rely=0.85,relwidth=0.11, relheight=0.05)
        self.encarregado_entry = ttk.Combobox(self.aba1,font=self.root.font,values=lista_encarregado)
        self.encarregado_entry.place(relx=0.15, rely=0.85, relwidth=0.27, relheight=0.06)


# Botao Buscar por MATERIAL
        self.bt_buscar = customtkinter.CTkButton(self.aba1, command=self.busca_mat,
                                                        text='Buscar por MATERIAL ', text_font=(self.root.font,9,'bold'), text_color="white",
                                                        width=900, compound=BOTTOM, relief=RAISED, 
                                                        hover= True,hover_color= "#00a9d4",
                                                        border_width=2, border_color= "lightblue", corner_radius=8,
                                                        fg_color= "#1c3166", bg_color='#2e97b7'
                                                )
        self.bt_buscar.place(relx=0.44, rely=0.68, relwidth=0.15, relheight=0.10)

# Botão Buscar por OS
        self.bt_busca_os = customtkinter.CTkButton(self.aba1, command=self.busca_os,
                                                        text='Buscar por OS', text_font=(self.root.font,9,'bold'), text_color="white",
                                                        width=900, compound=BOTTOM, relief=RAISED, 
                                                        hover= True, hover_color= "#00a9d4",
                                                        border_width=2,border_color= "lightblue",corner_radius=8,
                                                        fg_color= "#1c3166", bg_color='#2e97b7'
                                                )
        self.bt_busca_os.place(relx=0.44, rely=0.82, relwidth=0.15, relheight=0.10)

# Botão Busca Data
        self.calData = customtkinter.CTkButton(self.aba1,  command=self.busca_data,
                                                        text='Buscar por Data', text_font=(self.root.font,9,'bold'), text_color="white",
                                                        width=900, compound=BOTTOM, relief=RAISED, 
                                                        hover= True, hover_color= "#00a9d4",
                                                        border_width=2,border_color= "lightblue",corner_radius=8,
                                                        fg_color= "#1c3166", bg_color='#2e97b7',
                                                )
        self.calData.place(relx=0.44, rely=0.55, relwidth=0.15, relheight=0.09)
        
        self.bt_chama_calendario = customtkinter.CTkButton(self.aba1, command=self.calendario,
                                                        text='CALENDÁRIO', text_font=(self.root.font,9,'bold'), text_color="white",
                                                        width=900, compound=BOTTOM, relief=RAISED, 
                                                        hover= True,hover_color= "#00a9d4",
                                                        border_width=2,border_color= "lightblue",corner_radius=8,
                                                        fg_color= "#101652", bg_color='#2e97b7' 
                                                )
        self.bt_chama_calendario.place(relx=0.44, rely=0.34, relwidth=0.15, relheight=0.09)


        self.python_image2 = PhotoImage(data=base64.b64decode(self.somar))
        self.python_image2 = self.python_image2.subsample(2, 2)
        self.fechar2 = Label(self.frame_1, image=self.python_image2)
        self.fechar2.place(relx=0.62, rely=0.34, relwidth=0.37, relheight=0.63)


    def calendario(self):
        self.fechar2.place_forget()
        self.calendario1 = Calendar(self.aba1, fg='gray75', bg='blue', font=('Verdana', '9', 'bold'),
                                    locale='pt_br')
        self.calendario1.place(relx=0.62, rely=0.31, relwidth=0.37, relheight=0.65)
        
        self.calData = customtkinter.CTkButton(self.aba1, command=self.print_cal,
                                                        text='Inserir Data', text_font=(self.root.font,9,'bold'), text_color="white",
                                                        width=900, compound=BOTTOM, relief=RAISED, 
                                                        hover= True, hover_color= "#00a9d4",
                                                        border_width=2,border_color= "#0e002f",corner_radius=8,
                                                        fg_color= "#101652",bg_color='lightblue' 
                                                )
        self.calData.place(relx=0.62, rely=0.18, relwidth=0.10, relheight=0.10)


        # Insere o calendario no Campo

    def print_cal(self):
        dataini = self.calendario1.get_date()
        self.calendario1.destroy()
        self.data_entry.delete(0, END)
        self.data_entry.insert(END, dataini)
        self.calData.destroy()
        self.fechar2.place(relx=0.62, rely=0.34, relwidth=0.37, relheight=0.63)


# Imagem da Logo Frame 3, Canvas e Labels da Frame 2
    def imag(self):
        self.python_image = PhotoImage(data=base64.b64decode(self.lvl))
        self.python_image = self.python_image.subsample(3, 3)
        self.fechar = Label(self.frame_3, image=self.python_image)
        self.fechar.place(relx=0, rely=0, relwidth=0.99, relheight=0.99)



        self.canvas_tab_mat = Canvas(self.frame_2, bd=0, bg=self.root.cor_inter_aba3,
                                     highlightbackground=self.root.cor_bord_aba3, highlightthickness=5)
        self.canvas_tab_mat.place(relx=0, rely=0, relwidth=1, relheight=0.11)

        self.lb_tab_mat = Label(self.frame_2, text='Tabela da Saída do Material', font=('Verdana', 15, 'bold'), bg='#001f36',
                                fg='white')
        self.lb_tab_mat.place(relx=0.36, rely=0.02, relwidth=0.31, relheight=0.06)
# Canvas da TABELA
        self.canvas_bt = Canvas(self.frame_2, bd=3, bg='#2e97b7', highlightbackground='#001f36', highlightthickness=9)
        self.canvas_bt.place(relx=0, rely=0.12, relwidth=1, relheight=0.88)




# Montagem da Tabela da Frame 2
    def lista_frame2(self):
        self.lista_mat = ttk.Treeview(self.frame_2, height=3,columns=('col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7','col8'))
        self.lista_mat.heading('#0', text='')
        self.lista_mat.heading('#1', text='ID')
        self.lista_mat.heading('#2', text='DATA')
        self.lista_mat.heading('#3', text='ENDEREÇO')
        self.lista_mat.heading('#4', text='MATERIAL')
        self.lista_mat.heading('#5', text='QUANTIDADE')
        self.lista_mat.heading('#6',text='UNI. MEDIDA')
        self.lista_mat.heading('#7', text='ORD. SERVIÇO')
        self.lista_mat.heading('#8', text='ENCARREGADO')
        self.lista_mat.column('#0', width=1,anchor='center')
        self.lista_mat.column('#1', width=40,anchor='center')
        self.lista_mat.column('#2', width=50,anchor='center')
        self.lista_mat.column('#3', width=150,anchor='center')
        self.lista_mat.column('#4', width=150,anchor='center')
        self.lista_mat.column('#5', width=50,anchor='center')
        self.lista_mat.column('#6',width=50,anchor='center')
        self.lista_mat.column('#7', width=50,anchor='center')
        self.lista_mat.column('#8', width=100,anchor='center')
        self.lista_mat.place(relx=0.01, rely=0.25, relwidth=0.96, relheight=0.72) # Posição da Tabela

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.lista_mat.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.97,rely=0.25,relwidth=0.02,relheight=0.72)  # Posição da Scrollbar
        self.lista_mat.bind("<Double-1>",self.OnDoubleClick)

# Botão Gerar Excel
        self.gera_excel = Button(self.frame_2, text='Gerar PLANILHA', font=(self.root.font,7,'bold' ),bd=3, bg='#dce4f7',activebackground='#e30075', activeforeground='white', fg=self.root.font_cor,command=self.exportar_excel)
        self.gera_excel.place(relx=0.83, rely=0.16, relwidth=0.14, relheight=0.07)

        self.lb_plani = Label(self.frame_2,text='NOME DA PLANILHA', font=(self.root.font,9,'bold'), bd=5, bg='#2e97b7')
        self.lb_plani.place(relx=0.45, rely=0.16, relwidth=0.20, relheight=0.07)
# Botão Chama planilah Excel
        self.entry_excel = Entry(self.frame_2, font=(self.root.font,13,'bold'))
        self.entry_excel.place(relx=0.64, rely=0.16, relwidth=0.17, relheight=0.07)


# Tratamento de ERRO
    def validaEntrada(self):
        
        self.vcmd2 = (self.root.register(self.validate_entry2), "%P")
        
Application()
