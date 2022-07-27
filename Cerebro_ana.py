
import speech_recognition as sr
from gtts import gTTS
from playsound import playsound
import os
import pandas as pd

import teste


class cerebro_ana():
    def __init__(self,nome_user = "",nome_asistente = "Ana"):
        self.nome_asistente  = nome_asistente
        self.nome_user = nome_user
        self.microfone = sr.Recognizer()
        self.pandas = pd
        self.incializar_programa()
        self.iniciar = False
    def incializar_programa(self):
        self.comandos_para_ligar = ['ana iniciar','iniciar','iniciar ana']
        print("__" * 40)
        print("Ana Editora de planilhas".center(80))
        print("--" * 40)
        while (True):
            self.comando_iniciar = self.ouvindo_microfone_ana().lower()
            if self.comando_iniciar in self.comandos_para_ligar:
                while (True):
                    self.iniciar = True
                    self.ana_falando_class(audio="Qual é seu nome ? ", nome_audio="Ana_pedindo_nome_user")
                    self.nome_user = self.ouvindo_microfone_ana()
                    if (self.nome_user == ""):
                        True
                    else:
                        self.ana_falando_class(f"{self.nome_user} Meu nome anna, sou uma assistente virtual de edição de planilhas","Ana_se_apresentando")
                        self.abrir_pasta(self.nome_user)
            else:
                True
    def ouvindo_microfone_ana(self):
        self.contador = 0
        while(True):
            with sr.Microphone(device_index=0) as source:
                self.microfone.adjust_for_ambient_noise(source)
                self.audio = self.microfone.listen(source)
                try:
                    self.frase = self.microfone.recognize_google(self.audio, language='pt-BR')
                    print(f"{self.nome_user} >>> " + self.frase)
                except sr.UnknownValueError:
                    self.contador += 1
                    if(self.iniciar == True):
                        self.ana_falando_class(nome_audio="ana_nao_entedeu",audio="Não entendi,pode repetir")
                    else:
                        pass
                if(self.contador == 1):
                    True
                    self.contador -= 1
                elif(self.contador == 0):
                    return self.frase
    def abrir_pasta(self,nome_user):
        self.arquivo_nome = []
        while (True):
            self.arquivos_pasta = os.listdir("planilhas_ler")
            for i in range(len(self.arquivos_pasta)):
                self.ana_falando_class(nome_audio="falando_nome_pastas",audio=f"Numero {i} é {self.arquivos_pasta[i]}")
                print(f"Arquivo nº{i} = {self.arquivos_pasta[i]}")
            if (self.arquivos_pasta == ''):
                True
            self.ana_falando_class(f"{self.nome_user} Qual arquivo de planilhas o senhor quer abrir ","ana_pedindo_arquivo_para_editar")
            self.numero_arquivo = self.ouvindo_microfone_ana()
            self.numero_arquivo = int(self.numero_arquivo)
            self.ana_falando_class(f"{nome_user} Estou carregando arquivo {self.arquivos_pasta[self.numero_arquivo]} para edição","ana_aberindo_arquivo_para_editar")
            self.nome_arquivo_ = self.arquivos_pasta[self.numero_arquivo]
            self.editando_planilhas()
    def ana_falando_class(self,audio,nome_audio):
        self.tts = gTTS(text=audio, lang="pt-br")
        self.nome_audio = nome_audio
        self.audio_file = f'{self.nome_audio}.mp3'
        self.tts.save(self.audio_file)
        playsound(self.audio_file)
        print('Ana >>> ', audio )
        os.remove(self.audio_file)
    def editando_planilhas(self):
        self.pandas.options.display.max_rows = 0
        while(True):
            print('__' * 40)
            print("Vamos começar editar".center(80))
            print('--' * 40)
            if("csv" in self.nome_arquivo_):
                self.nome_arquivo_csv = self.nome_arquivo_
                self.editando_csv = self.pandas.read_csv(f"planilhas_ler\\{self.nome_arquivo_csv}", sep=';', encoding='latin-1')
                print(self.editando_csv.to_string())
                self.editando_iniciar(self.editando_csv)
            elif("xls" in self.nome_arquivo_):
                self.nome_arquivo_xlsx = self.nome_arquivo_
                self.editando_xlsx = self.pandas.read_excel(f"planilhas_ler\\{self.nome_arquivo_xlsx}",index_col=0)
                print(self.editando_xlsx)
                self.editando_iniciar(self.editando_xlsx)
    def editando_iniciar(self,name_planilha):
        while(True):
            self.editando_essa_planilha = name_planilha
            self.lista_nomes_coluna = self.editando_essa_planilha.columns.str.lower()
            print("__" * 50)
            print("COMANDOS".center(90))
            print("--" * 50)
            print("[DELETAR] - [ADICIONAR] - [MOSTRA-PLANILHA] - [ABRIR-OUTRA-PLANILHA] - [SAIR]".center(90))
            self.ana_falando_class(nome_audio="ana_falando_comandos",audio=f"{self.nome_user} Esses são os comandos , deletar , adicionar , mostrar planilha ,abrir outra planilha  é sair ")
            self.ana_falando_class(nome_audio="ana_falando_comandos_qual",audio=f"{self.nome_user} Fale nome do comando que deseja aplicar na planilha ")
            self.comando_executar = self.ouvindo_microfone_ana().lower()
            self.lista_comandos_editar = ["deletar","adicionar","mostrar planilha","abrir outra planilha","sair"]
            if(self.comando_executar not in self.lista_comandos_editar):
                True
            else:
                if(self.comando_executar == self.lista_comandos_editar[0]):
                    print(self.editando_essa_planilha.columns)
                    while(True):
                        self.comandos_deletar = ["deletar coluna", "deletar linha da coluna","deletar intem da coluna", "sair"]
                        print("[DELETAR-COLUNA] - [DELETAR-LINHA-DA-COLUNA] - [DELETAR-INTEM-DA-COLUNA] - [SAIR]")
                        self.ana_falando_class(nome_audio="ana_perguntando_comando",audio=f"Temos quatro comandos em deletar {self.nome_user} \n"f"{self.comandos_deletar}")
                        self.comando_executar_deletar = self.ouvindo_microfone_ana().lower()
                        if(self.comando_executar_deletar not  in self.comandos_deletar):
                            self.ana_falando_class(nome_audio="ana_comando_existenti",audio=f"{self.nome_user} comando não existe")
                            True
                        else:
                            if(self.comando_executar_deletar == self.comandos_deletar[0]):
                                while(True):
                                    self.ana_falando_class(nome_audio="ana_deletando_coluna",audio=f"{self.nome_user} qual nome da coluna que você deseja excluir")
                                    print(self.editando_essa_planilha.columns)
                                    self.executar_comando = self.ouvindo_microfone_ana().lower()
                                    try:
                                        self.editando_essa_planilha = self.editando_essa_planilha.drop(self.executar_comando,axis=1)
                                        self.ana_falando_class(nome_audio="ana_deletou",audio=f"{self.nome_user} deletei a coluna {self.executar_comando}")
                                        if('csv' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_csv(f'planilhas_ler\\{self.nome_arquivo_}')
                                        elif('xls' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_excel(f'planilhas_ler\\{self.nome_arquivo_}')
                                    except:
                                        True
                            if(self.comando_executar == self.comandos_deletar[1]):
                                while(True):
                                    self.ana_falando_class(nome_audio="ana_deletando_linha",audio=f"{self.nome_user} Qual index da linha que você deseja excluir")
                                    self.executar_comando = self.ouvindo_microfone_ana()
                                    try:
                                        self.editando_essa_planilha.drop(f'{self.executar_comando}', axis=0)
                                        self.ana_falando_class(nome_audio="ana_deletou",audio=f"{self.nome_user} deletei a linha {self.executar_comando}")
                                        if ('csv' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_csv(f'planilhas_ler\\{self.nome_arquivo_}')
                                        elif ('xls' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_excel(f'planilhas_ler\\{self.nome_arquivo_}')
                                    except:
                                        True
                            if (self.comando_executar == self.comandos_deletar[2]):
                                while(True):
                                    self.ana_falando_class(nome_audio="ana_deletando_intem",audio=f"{self.nome_user} Qual nome da coluna do intem você deseja excluir")
                                    self.executar_comando = self.ouvindo_microfone_ana()
                                    if(self.comando_executar  not  in self.comandos_deletar):
                                        self.ana_falando_class(nome_audio="ana_nao_existi",audio="nome de coluna não existe")
                                        True
                                    self.ana_falando_class(nome_audio="ana_nome_deletar",audio="nome do intem para deletar")
                                    self.comando_executar_intem = self.ouvindo_microfone_ana()
                                    try:
                                        self.editando_essa_planilha = self.editando_essa_planilha[self.comando_executar].drop(f'{self.comando_executar_intem}', axis=0)
                                        self.ana_falando_class(nome_audio="ana_deletou",audio=f"{self.nome_user} deletei o intem {self.executar_comando}")
                                        if ('csv' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_csv(f'planilhas_ler\\{self.nome_arquivo_}')
                                        elif ('xls' in self.nome_arquivo_):
                                            self.editando_essa_planilha.to_excel(f'planilhas_ler\\{self.nome_arquivo_}')
                                    except:
                                        True
                            if(self.comando_executar == self.comandos_deletar[3]):
                                self.ana_falando_class(nome_audio="ana_saindo_funcao",audio=f"Saindo da funão {self.nome_user}")
                                break
                elif(self.comando_executar == self.lista_comandos_editar[1]):
                    while(True):
                        self.comandos_adicionar = ["trocar nome da coluna","trocar nome de item na coluna","adicionar intem a coluna","sair"]
                        print("[TROCAR-NOME-DA-COLUNA] - [TROCAR NOME DE INTEM NA COLUNA][ADICIONAR-INTEM-A-COLUNA] - [SAIR]")
                        self.ana_falando_class(nome_audio="ana_perguntando_comando",audio=f"Temos dois comandos em adicionar {self.nome_user} \n"
                                                                                          f"trocar nome da coluna , adicionar intem a coluna é sair")
                        self.comando_executar_adicionar = self.ouvindo_microfone_ana()
                        if self.comando_executar_adicionar not in self.comandos_adicionar:
                            self.ana_falando_class(nome_audio="ana_falando_comando_nao_existe",audio=f"Desculpa {self.nome_user} esse comando não existe")
                            True
                        elif self.comando_executar_adicionar == self.comandos_adicionar[0]:
                            while(True):
                                print(self.lista_nomes_coluna)
                                self.ana_falando_class(nome_audio="ana_qual_nome_coluna",audio="Qual é nome da coluna que você deseja altera")
                                self.ana_falando_class(nome_audio="ana_qual_executar",audio=f"{self.lista_nomes_coluna}")
                                self.nome_da_coluna = self.ouvindo_microfone_ana().lower()
                                if(self.nome_da_coluna not in self.lista_nomes_coluna):
                                    self.ana_falando_class(nome_audio="ana_dizendo",audio="Esse nome não e nome de colunas dessa planilha")
                                    True
                                self.ana_falando_class(nome_audio="ana_qual_nome_novo",audio="Qual é novo nome da coluna")
                                self.nome_da_coluna_novo = self.ouvindo_microfone_ana()
                                if (self.nome_da_coluna_novo not in self.lista_nomes_coluna):
                                    self.ana_falando_class(nome_audio="ana_dizendo",audio="Esse nome ja existe na planilha")
                                    True
                                self.editando_essa_planilha = self.editando_essa_planilha.columns.str.lower()
                                try:
                                    self.editando_essa_planilha.rename(columns={f"{self.nome_da_coluna}":f"{self.nome_da_coluna_novo}"})
                                    self.ana_falando_class(nome_audio="ana_trocamos",audio=f"{self.nome_user} Trocamos {self.nome_da_coluna} por {self.nome_da_coluna_novo}")
                                    if ('csv' in self.nome_arquivo_):
                                        self.editando_essa_planilha.to_csv(f'planilhas_ler\\{self.nome_arquivo_}')
                                    elif ('xls' in self.nome_arquivo_):
                                        self.editando_essa_planilha.to_excel(f'planilhas_ler\\{self.nome_arquivo_}')
                                        break
                                except:
                                    True
                                break
                        elif(self.comando_executar_adicionar == self.comandos_adicionar[1]):
                            while (True):
                                print(self.lista_nomes_coluna)
                                self.ana_falando_class(nome_audio="ana_qual_nome_coluna",audio="Qual é nome da coluna que você deseja mudar intem dela")
                                self.ana_falando_class(nome_audio="ana_qual_executar",audio=f"{self.lista_nomes_coluna}")
                                self.nome_da_coluna = self.ouvindo_microfone_ana().lower()
                                if (self.nome_da_coluna not in self.lista_nomes_coluna):
                                    self.ana_falando_class(nome_audio="ana_dizendo",audio="Esse nome não e nome de uma coluna dessa planilha")
                                    True
                                self.ana_falando_class(nome_audio="ana_qual_nome_novo",audio="Qual valor que deseja altera ")
                                self.nome_alterar = self.ouvindo_microfone_ana()
                                self.ana_falando_class(nome_audio="ana_qual_nome_novo",audio="Qual é novo valor ")
                                self.nome_alterar_novo = self.ouvindo_microfone_ana()
                                try:
                                    self.editando_essa_planilha[self.nome_da_coluna].replace(f"{self.nome_alterar}", f"{self.nome_alterar_novo}")
                                    self.ana_falando_class(nome_audio="ana_trocamos",audio=f"{self.nome_user} Trocamos {self.nome_alterar} por {self.nome_alterar_novo}")
                                    if ('csv' in self.nome_arquivo_):
                                        self.editando_essa_planilha.to_csv(f'planilhas_ler\\{self.nome_arquivo_}')
                                    elif ('xls' in self.nome_arquivo_):
                                        self.editando_essa_planilha.to_excel(f'planilhas_ler\\{self.nome_arquivo_}')
                                        break
                                except:
                                    True
                                break
                        elif(self.comando_executar_adicionar == self.comandos_adicionar[2]):
                            while (True):
                                print(self.lista_nomes_coluna)
                                self.ana_falando_class(nome_audio="ana_qual_nome_coluna",audio="Qual é nome da coluna que você deseja mudar intem dela")
                                self.ana_falando_class(nome_audio="ana_qual_executar",audio=f"{self.lista_nomes_coluna}")
                                self.nome_da_coluna = self.ouvindo_microfone_ana().lower()
                                self.ana_falando_class(nome_audio='ana_nada_ver',audio='Desculpas nâo fui programada para adicionar ainda')
                            break
                        elif(self.comando_executar_adicionar == self.comandos_adicionar[3]):
                            self.ana_falando_class(nome_audio="ana_saindo_funcao",audio=f"Saindo da funão {self.nome_user}")
                            break
                elif(self.comando_executar == self.lista_comandos_editar[2]):
                    while(True):
                        print('[informações da planilha] - [Contagem de dados não nulos] - [Nome das colunas] - [Descrição da planilha] - [Quantidade de linhas e colunas] - [Sair]')
                        self.comandos_mostra_planilhas = ['informações da planilha','contagem de dados não nulos','nome das colunas',
                                                          'descrição da planilha','quantidade de linhas e colunas','Sair']
                        self.ana_falando_class(nome_audio="ana_falando_comandos",audio=f"{self.nome_user} Esses são os comandos {self.comandos_mostra_planilhas}")
                        self.ana_falando_class(nome_audio="ana_falando_comandos_qual",audio=f"{self.nome_user} Fale nome do comando que deseja aplicar na planilha ")
                        self.comando_executar_mostra_planilhas = self.ouvindo_microfone_ana().lower()
                        if(self.comando_executar_mostra_planilhas not in self.comandos_mostra_planilhas):
                            self.ana_falando_class(nome_audio="ana_falando_comandos",audio=f"{self.nome_user} Esse Comando não existe")
                            True
                        else:
                            if(self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[0]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'{self.editando_essa_planilha.info()}')
                            elif(self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[1]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'{self.editando_essa_planilha.count()}')
                            elif (self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[2]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'{self.editando_essa_planilha.columns()}')
                            elif (self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[3]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'{self.editando_essa_planilha.describe()}')
                            elif (self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[4]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'{self.editando_essa_planilha.shape()}')
                            elif (self.comando_executar_mostra_planilhas == self.comandos_mostra_planilhas[5]):
                                self.ana_falando_class(nome_audio='ana_falando',audio=f'Obrigada saindo da função mostrando planilha')
                                break
                elif(self.comando_executar == self.lista_comandos_editar[3]):
                    self.ana_falando_class(nome_audio="ana_abrindo_nova_planilha",audio=f"{self.nome_user} Estou abrindo a pasta de planilhas para você")
                    self.abrir_pasta(self.nome_user)
                elif(self.comando_executar == self.lista_comandos_editar[5]):
                    self.ana_falando_class(nome_audio="ana_desligando",audio=f"obrigado {self.nome_user}, gostou da experiencia")
                    exit()


