from models.projeto import Dia, Semana, Projeto

import win32com.client as win32
from datetime import timezone, date, timedelta
import shutil
import os
import pywintypes
import streamlit as st

class WorkbookManager:
    def __init__(self, filename):
        self.original_filename = filename
        self.copied_filename = self.create_copy(filename)
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.workbook = self.excel.Workbooks.Open(self.copied_filename)
        self.cronograma_geral = CronogramaGeral(self.workbook)
        self.proximas_semanas = ProximasSemanas(self.workbook)
        self.planilhas_semanais = PlanilhasSemanais(self.workbook)

    def create_copy(self, filename):
        # Cria o nome do arquivo de cópia com o novo nome desejado
        new_filename = os.path.join(os.path.dirname(filename), "Master Planejamento Atualizado.xlsx")
        
        # Copia o arquivo original para o novo arquivo
        shutil.copy(filename, new_filename)
        
        # Retorna o caminho completo do novo arquivo
        return os.path.abspath(new_filename)
    
    def save(self):
        # Salva a cópia alterada
        self.workbook.Save()
        self.excel.Quit()

    def close(self):
        self.excel.Quit()


class CronogramaGeral:

    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = self.workbook.Sheets('Crono. Geral')
        
        # Atributos posicionais da planilha
        self.lin_datas = 6
        self.lin_inicio_lista_projetos = 7
        self.col_inicio_datas = 10
        self.col_nomes_projetos = 2
        self.lin_feriado = 5


    def checa_feriado(self, col):
        '''
        Função booleana que checa se uma data é feriado

        Parameters
        ----------
        col (int) -> Número da coluna que se encontra a data a ser checada
        '''
        if self.sheet.Cells(self.lin_feriado, col).Value == "FERIADO":
            return True
        else:
            return False


    def get_data_inicio(self, data_inicio_coleta, dias_pre):
        """
        Retorna a data de início de um projeto considerando os dias pré coleta
        """
        coluna = self.get_coluna_inicio(data_inicio_coleta)
        linha_datas = self.lin_datas
        dias = dias_pre
        
        while dias > 0:
            
            if not self.checa_feriado(coluna):
                dias -= 1
                
            coluna -= 1
            
        data = self.sheet.Cells(linha_datas, coluna).Value.date()
        return data


    def get_coluna_inicio(self, data_inicio):
        '''
        Retorna o número da coluna de início do projeto
        '''
        col = self.col_inicio_datas
        while True:
            # data_inicio = data_inicio.replace(tzinfo=timezone.utc)
            data_plan = self.sheet.Cells(self.lin_datas, col).Value.date()
            if data_plan == data_inicio:
                return col
            col += 1
            if col > self.sheet.Cells(self.lin_datas, self.sheet.Columns.Count).End(-4161).Column:  # -4161 corresponde a xlToLeft
                break
        return None


    def get_data_fim(self, data_inicio, dias_uteis):
        '''
        Retorna o dia final do projeto
        '''
        col_inicio_proj = self.get_coluna_inicio(data_inicio)
        if col_inicio_proj is None:
            print("Data de início não encontrada.")
            return None

        i = 0
        col_fim = col_inicio_proj
        while i < dias_uteis:
            if not self.checa_feriado(col_fim):
                i += 1
            col_fim += 1

        # Adaptando para encontrar a data final usando pywin32
        data_fim = self.sheet.Cells(self.lin_datas, col_fim - 1).Value
        return data_fim

    
    def get_linha_crono_geral(self, col_inicio_proj):
        '''
        Retorna o número da linha em que o projeto será colocado na planilha
        '''
        lin_inicio_lista_projetos = self.lin_inicio_lista_projetos
        col_inicio_datas = self.col_inicio_datas
        
        # Acessa a última linha com dados na planilha para definir um limite de busca
        ultima_linha_com_dados = self.sheet.Cells(self.sheet.Rows.Count, col_inicio_proj).End(win32.constants.xlUp).Row
    
        for lin in range(lin_inicio_lista_projetos, ultima_linha_com_dados + 1):
            celula_preenchida = False
            for col in range(col_inicio_proj, col_inicio_datas - 1, -1):
                # Verifica se a célula nesta posição está preenchida
                if self.sheet.Cells(lin, col).Value is not None:
                    celula_preenchida = True
                    break
            if not celula_preenchida:
                return lin
        return ultima_linha_com_dados + 1  # Retorna a próxima linha vazia se não encontrar uma linha adequada


    def inserir_projeto(self, specifications, projeto):
        '''
        Insere as informações na planilha.
        '''
        data_inicio = self.get_data_inicio(projeto.data_inicio_coleta, 5)
        col_inicio_proj = self.get_coluna_inicio(data_inicio)

        if col_inicio_proj is None:
            print("Coluna de início não encontrada.")
            return
        
        lin_posicao_projeto = self.get_linha_crono_geral(col_inicio_proj)

        rang = self.sheet.Range(f"A{lin_posicao_projeto}")
        rang.EntireRow.Insert(Shift=win32.constants.xlShiftDown, CopyOrigin=win32.constants.xlFormatFromLeftOrAbove)

        # Configurando o nome do projeto na posição especificada
        self.sheet.Cells(lin_posicao_projeto, self.col_nomes_projetos).Value = projeto.nome
        self.sheet.Cells(lin_posicao_projeto, self.col_nomes_projetos).Font.Name = 'Verdana'
        self.sheet.Cells(lin_posicao_projeto, self.col_nomes_projetos).Font.Size = 11

        col = col_inicio_proj
        for text, count in specifications:
            while count > 0:
                if not self.checa_feriado(col):
                    cell = self.sheet.Cells(lin_posicao_projeto, col)
                    cell.Value = text
                    count -= 1
                col += 1
                
                
    # def get_data_inicio_coleta(self, projeto):
        
    #     inicio = self.get_coluna_inicio(projeto.data_inicio)
    #     fim = inicio + projeto.duracao_dias
    #     linha = self.get_linha_crono_geral(inicio)
        
    #     for col in range(inicio, fim):
    #         if self.sheet.Cells(linha, col).Value == "C":
    #             col_inicio_coleta = col
                
    #     data = self.sheet.Cells(self.lin_datas, col_inicio_coleta).Value
        
    #     return data
    

class ProximasSemanas:
    
    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = self.workbook.Sheets('Próximas Semanas')
        
        # Atributos posicionais da planilha
        self.lin_inicio_lista_projetos = 6
        self.col_PB = 3
        self.col_data_inicio = 6
        self.col_data_fim = 7
        self.col_amostra = 8
        self.col_coleta = 9
        self.col_dur = 11
        self.col_media = 12
        self.col_tma = 13
        self.col_nomes = 16
        
        
    def get_linha_projeto(self, data_inicio_coleta):
        '''
        Retorna o número da linha em que o projeto será colocado na planilha
        '''
        # data_inicio_coleta = data_inicio_coleta.replace(tzinfo=timezone.utc);
        
        last_row = self.sheet.UsedRange.Rows.Count
        
        for lin in range(self.lin_inicio_lista_projetos, last_row + 1):
            data_celula = self.sheet.Cells(lin, self.col_data_inicio).Value.date()
            if data_celula >= data_inicio_coleta:
                linha_projeto = lin
                break
            
        return linha_projeto
    
    
    def get_diff_producao_pessoas(self):
        """
        Retorna a diferença de linhas da posição de um projeto na tabela
        de produção para esse mesmo projeto na tabela de pessoas.
        """
        last_row = self.sheet.UsedRange.Rows.Count
        linha_producao = None
        linha_pessoas = None
        
        for i in range(1, last_row + 1):
            valor_celula = self.sheet.Cells(i, self.col_nomes).Value
            if valor_celula == "Produção":
                linha_producao = i
            elif valor_celula == "Pessoas":
                linha_pessoas = i
                
            # Se ambos foram encontrados, não precisa continuar
            if linha_producao is not None and linha_pessoas is not None:
                break
    
        return abs(linha_pessoas - linha_producao)
    
    
    def get_linha_pessoas(self, data_inicio_coleta):
        """
        Retorna a posição do projeto na tabela de pessoas.
        """
        return self.lin_inicio_lista_projetos + self.get_diff_producao_pessoas()
    
        
    def inserir_projeto_prod(self, projeto):
        '''
        Insere as informações na tabela de produção.
        '''  
        # mudar isso aqui
        media = 1
        tma = 1
        
        # TABELA COLETA
        linha = self.get_linha_projeto(projeto.data_inicio_coleta)
        
        range = self.sheet.Rows(linha)
        # Copia a linha acima para manter as fórmulas
        range.Copy()
        self.sheet.Rows(linha + 1).Insert(Shift=win32.constants.xlDown)
        self.sheet.Rows(linha + 1).PasteSpecial(Paste=win32.constants.xlPasteFormulas)
        
        # Adiciona infos do projeto
        self.sheet.Cells(linha, self.col_nomes).Value = projeto.nome
        self.sheet.Cells(linha, self.col_data_inicio).Value = pywintypes.Time(projeto.data_inicio_coleta)
        self.sheet.Cells(linha, self.col_data_fim).Value = pywintypes.Time(projeto.data_fim)
        self.sheet.Cells(linha, self.col_amostra).Value = projeto.amostra
        self.sheet.Cells(linha, self.col_dur).Value = projeto.duracao_coleta
        self.sheet.Cells(linha, self.col_media).Value = media
        self.sheet.Cells(linha, self.col_tma).Value = tma
        
        
    def inserir_projeto_pessoas(self, projeto):
        """
        Insere as informações na tabela de pessoas.
        """
        # TABELA PESSOAS
        linha = self.get_linha_pessoas(projeto.data_inicio_coleta)
        
        range = self.sheet.Rows(linha)
        # Copia a linha acima para manter as fórmulas
        range.Copy()
        self.sheet.Rows(linha + 1).Insert(Shift=win32.constants.xlDown)
        self.sheet.Rows(linha + 1).PasteSpecial(Paste=win32.constants.xlPasteFormulas)
        
        # Adiciona infos do projeto
        self.sheet.Cells(linha, self.col_nomes).Value = projeto.nome
        
        
    def inserir_projeto(self, projeto):
        """
        Insere todas as informações do projeto na planilha
        """
        self.inserir_projeto_prod(projeto)
        self.inserir_projeto_pessoas(projeto)
        

class PlanilhasSemanais:
    
    def __init__(self, workbook):
        self.workbook = workbook
        
        # Colunas
        self.col_nomes = 5
        self.col_dias_inicio = 6
        self.col_coleta_real = 8
        self.col_hc_real = 11
        self.col_coleta_total = 15
        self.col_prod = 16
        self.col_hc_total = 17
        self.col_hc_parcial = 20
        self.col_dia = self.get_coluna_diaria()
        
        # Linhas
        self.lin_nomes_inicio = 8
        
    def get_sheets(self, projeto):
        """
        Pega as planilhas semanais em que determinado projeto se encontra.
        """
        sheets = []
        for semana in projeto.semanas:
            data_str = semana.data_inicio.strftime('%d.%m')
            sheetname = "CATI_Semana_" + data_str
            sheet = self.workbook.Sheets(sheetname)
            sheets.append(sheet)
            
        return sheets
    
    
    def get_lin_monitoramento(self, sheet):
        """
        Pega a linha que começa a tabela de monitoramento - varia a cada planilha.
        """
        linha = 1
        while True:
            if sheet.Cells(linha, self.col_nomes).Value == "Monitoramento":
                return linha + 3
            linha += 1
            
        
    def inserir_projeto(self, projeto):
        """
        Insere o projeto nas planilhas semanais corretas.
        """
        linha = self.lin_nomes_inicio
        sheets = self.get_sheets(projeto)
        for sheet, semana in zip(sheets, projeto.semanas):
            
            # ADICIONANDO NA TABELA DE DIMENSIONAMENTO
            
            range = sheet.Rows(linha)
            # Copia a linha acima para manter as fórmulas
            range.Copy()
            sheet.Rows(linha + 1).Insert(Shift=win32.constants.xlDown)
            sheet.Rows(linha + 1).PasteSpecial(Paste=win32.constants.xlPasteFormulas)
            
            sheet.Cells(linha, self.col_nomes).Value = projeto.nome
            i=0
            for dia in semana.dias_uteis:
                if not dia.feriado:
                    sheet.Cells(linha, self.col_dias_inicio + i).Value = dia.hc
                i += 1
            sheet.Cells(linha, self.col_prod).Value = semana.produtividade
            
            # ADICIONANDO NA TABELA DE MONITORAMENTO
            
            linha_monitoramento = self.get_lin_monitoramento(sheet)
            
            range = sheet.Rows(linha_monitoramento)
            # Copia a linha acima para manter as fórmulas
            range.Copy()
            sheet.Rows(linha_monitoramento + 1).Insert(Shift=win32.constants.xlDown)
            sheet.Rows(linha_monitoramento + 1).PasteSpecial(Paste=win32.constants.xlPasteFormulas)
            
            sheet.Cells(linha_monitoramento, self.col_nomes).Value = projeto.nome


    def get_planilha_semana_atual(self):
        """
        Retorna o nome da planilha baseada no dia de hoje.
        """
        hoje = date.today()
        dia_da_semana = hoje.weekday()
        
        inicio_semana = hoje - timedelta(days=dia_da_semana)
        str_data = inicio_semana.strftime('%d.%m')
        
        nome_planilha = "CATI_Semana_" + str_data
        
        return nome_planilha
    
    
    def get_coluna_diaria(self):
        """
        Retorna a coluna correspondente ao dia de hoje.
        """
        dia = date.today().weekday()
        
        if dia <= 5:
            return dia + self.col_nomes + 1
        
        # Exceção: domingo (fica na mesma coluna que sábado)
        if dia == 6:
            return dia + self.col_nomes
    
    
    def get_projetos_disponiveis(self):
        """
        Retorna os projetos disponíveis na planilha dessa semana.
        """
        projetos = []
        
        start_row = self.lin_nomes_inicio
        sheetname = self.get_planilha_semana_atual()
        sheet = self.workbook.Sheets(sheetname)
        
        while True:
            nome_projeto = sheet.Cells(start_row, self.col_nomes).Value
            valor_hc = sheet.Cells(start_row, self.col_dia).Value
            valor_coleta = sheet.Cells(start_row, self.col_coleta_total).Value
    
            # Se a célula em self.col_nomes for vazia, interrompe o loop
            if nome_projeto is None:
                break
    
            # Só adiciona o projeto se col_coleta_total não estiver vazia e col_dia for > 0
            if valor_coleta is not None and isinstance(valor_hc, (int, float)) and valor_hc > 0:
                projetos.append(nome_projeto)
    
            # Incrementa para mover para a próxima linha
            start_row += 1
    
        return list(set(projetos))
    
    
    def atualizar_coleta_diaria(self, coletas_por_projeto, hc_por_projeto):
        """
        Atualiza a planilha de coletas dessa semana de acordo com os inputs do usuário
        """
        
        sheetname = self.get_planilha_semana_atual()
        sheet = self.workbook.Sheets(sheetname)

        linha_atual = self.get_lin_monitoramento(sheet)
        
        while True:
            # Lê o nome do projeto na linha atual
            nome_projeto = sheet.Cells(linha_atual, self.col_nomes).Value
        
            # Se a célula estiver vazia, significa que não há mais projetos listados
            if nome_projeto is None:
                break
        
            # Se o projeto estiver no dicionário de coletas, adicione a coleta correspondente na coluna 'col_coleta_real'
            if nome_projeto in coletas_por_projeto:
                sheet.Cells(linha_atual, self.col_coleta_real).Value = coletas_por_projeto[nome_projeto]
                sheet.Cells(linha_atual, self.col_hc_real).Value = hc_por_projeto[nome_projeto]
            # Incrementa a linha para continuar para o próximo projeto
            linha_atual += 1
            
    def atualizar_meta_parcial(self):
        
        # Mapeia os dias da semana para as colunas correspondentes que serão somadas
        dia_para_colunas = {
            0: ['F'],       # Segunda-feira
            1: ['F', 'G'],  # Terça-feira
            2: ['F', 'G', 'H'],  # Quarta-feira
            3: ['F', 'G', 'H', 'I'],  # Quinta-feira
            4: ['F', 'G', 'H', 'I', 'J']  # Sexta-feira
        }
        
        # Pega o dia da semana atual (0=Segunda, 1=Terça, ..., 6=Domingo)
        hoje = date.today().weekday()
        
        # Acessa a planilha atual
        sheet = self.workbook.Sheets(self.get_planilha_semana_atual())
        
        # Inicia na linha definida por lin_nomes_inicio e continua até que uma célula vazia seja encontrada
        linha_atual = self.lin_nomes_inicio
        
        while True:
            # Verifica se chegou em uma célula vazia na coluna dos nomes dos projetos, se sim, para o loop
            if sheet.Cells(linha_atual, self.col_nomes).Value is None:
                break
        
            # Define a fórmula de acordo com o dia da semana
            if hoje in dia_para_colunas:
                colunas = dia_para_colunas[hoje]
                formula = f"=SUM({','.join(sheet.Cells(linha_atual, ord(col) - ord('A') + 1).Address for col in colunas)})"
                sheet.Cells(linha_atual, self.col_hc_parcial).Formula = formula
        
            # Incrementa a linha
            linha_atual += 1
        