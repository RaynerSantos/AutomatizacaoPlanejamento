from datetime import timedelta
import holidays

feriados = holidays.country_holidays("BR", subdiv="MG")
feriados_2024 = feriados["2024-01-01":"2024-12-31"]
feriados_2025 = feriados["2025-01-01":"2025-12-31"]

class Dia:
    
    def __init__(self, data):
        self.data = data
        self.hc = 0
        self.weekday = data.weekday()
        self.feriado = self.isFeriado()
        
    
    def isFeriado(self):
        
        data = self.data
        if data in feriados_2025:
            return True
        
        else:
            return False
        

class Semana:
    
    def __init__(self, data_inicio):
        self.data_inicio = data_inicio
        self.produtividade = 0 # Coluna P
        self.dias_uteis = [Dia(data_inicio + timedelta(days=i)) for i in range(5)]


    def get_coleta_semanal(self):
        """
        Retorna a coleta total planejada da semana para um projeto.
        Coluna O das planilhas semanais.
        """
        prod = self.produtividade
        hc = 0
        for dia in self.dias_uteis:
            hc += dia.hc
            
        return hc * prod
    
    
    def get_hc_semanal(self):
        """
        Retorna o hc total planejado da semana para um projeto.
        Coluna Q das planilhas semanais.
        """
        hc = 0
        for dia in self.dias_uteis:
            hc += dia.hc

        return hc
    
    
    def get_meta_parcial_hc(self, dia_meta):
        """
        Retorna a meta parcial do HC, dado um dia da semana.
        Coluna T das planilhas semanais.
        """
        meta = 0
        weekday = dia_meta.weekday
        for i in range(weekday):
            meta += self.dias_uteis[i].hc
    
        return meta
    
    
    def get_meta_parcial_coleta(self, dia_meta):
        """
        Retorna a meta parcial da coleta, dado um dia da semana.
        Coluna S das planilhas semanais.
        """
        meta_hc = self.get_meta_parcial_hc(dia_meta)
        
        return meta_hc * self.produtividade
    

class Projeto:
    
    def __init__(self, nome, amostra, data_inicio_coleta, duracao_coleta):
        self.nome = nome
        self.amostra = amostra
        self.data_inicio_coleta = data_inicio_coleta
        self.duracao_coleta = duracao_coleta
        # self.data_inicio_prep = self.get_data_inicio_prep()
        self.data_fim = self.get_data_fim()
        self.semanas = self.get_semanas()
        self.dias = self.get_dias()
    
    
    def get_data_inicio_semana(self):
        """
        Ajusta a data de início para a segunda-feira mais próxima no passado.
        """
        weekday = self.data_inicio_coleta.weekday()
        # Segunda-feira é 0 e domingo é 6, subtrai-se o dia da semana atual para voltar à segunda-feira
        return self.data_inicio_coleta - timedelta(days=weekday)
    
    
    def get_semanas(self):
        """
        Retorna as semanas em que a coleta acontecerá
        """
        semanas = []
        # Adiciona a primeira semana
        inicio_primeira_semana = self.get_data_inicio_semana()
        semanas.append(Semana(inicio_primeira_semana))
        
        # Loop pelos dias (começando no segundo dia pois o primeiro já foi adicionado)
        dia = self.data_inicio_coleta + timedelta(days=1)
        i = 0
    
        while i < self.duracao_coleta:
            # Apenas conta dias úteis
            if not (dia.weekday() == 5 or dia.weekday() == 6):
                i += 1
            # Se entrar em uma nova semana (segunda-feira), adiciona ela
            if dia.weekday() == 0:
                semanas.append(Semana(dia))
            # Atualiza a data
            dia += timedelta(days=1)
    
        return semanas
    
    
    def get_dias(self):
        """
        Retorna a lista de Dias do projeto
        """
        data = self.data_inicio_coleta
        dias = []

        i = 0
        while i < self.duracao_coleta:
            # Apenas conta dias úteis
            dia = Dia(data)
            if not (data.weekday() == 5 or data.weekday() == 6 or dia.feriado):
                dias.append(dia)
                i += 1
            # Atualiza a data  
            data += timedelta(days=1)
            
        return dias
    
    
    def get_data_fim(self):
        '''
        Retorna a data de fim do projeto, com base na data de início e na duração

        '''
        data = self.data_inicio_coleta
        
        i = 0
        while i < self.duracao_coleta - 1:
            # Apenas conta dias úteis
            dia = Dia(data)
            if not (data.weekday() == 5 or data.weekday() == 6 or dia.feriado):
                i += 1
            # Atualiza a data  
            data += timedelta(days=1)
            
        while (data.weekday() == 5 or data.weekday() == 6 or dia.feriado):
            data += timedelta(days=1)
            
        return data
    
    
    def get_coleta_total(self):
        """
        Retorna a coleta total (planejada) do projeto
        """
        coleta = 0
        for semana in self.semanas:
            coleta += semana.get_coleta_semanal()
            
        return coleta


    def get_gap(self):
        """
        Retorna o saldo de entrevistas (coleta planejada - amostra)
        Coluna B da planilha Próximas Semanas.
        """
        return self.get_coleta_total() - self.amostra
    
    
    # def get_data_inicio_prep(self):
        
    #     dia = self.data_inicio_coleta
    #     i = 0
    
    #     while i < self.duracao_pre_coleta - 1:
    #         # Apenas conta dias úteis
    #         if not (dia.weekday() == 5 or dia.weekday() == 6):
    #             i += 1
    #         # Se entrar em uma nova semana (segunda-feira), adiciona ela
    #         # Atualiza a data
    #         dia -= timedelta(days=1)
            
    #     return dia
    
    def input_info(self):
        """
        Pega as informações do projeto a partir do usuário
        """
        curr_semana = 1
        for semana in self.semanas:
            semana.produtividade = int(input(f"Produtividade da semana {curr_semana}: "))
            curr_semana += 1
            for dia in semana.dias_uteis:
                if dia.data >= self.data_inicio_coleta and dia.data <= self.data_fim:
                    dia.hc = int(input("HC do dia " + dia.data.strftime("%Y-%m-%d") + ": "))
                    
                    
    def input_info_hc_fixo(self, hc):
        """
        Pega as informações do projeto a partir do usuário - quando o projeto tem HC fixo
        """
        curr_semana = 1
        for semana in self.semanas:
            semana.produtividade = int(input(f"Produtividade da semana {curr_semana}: "))
            curr_semana += 1
            for dia in semana.dias_uteis:
                if dia.data >= self.data_inicio_coleta and dia.data <= self.data_fim:
                    dia.hc = hc



