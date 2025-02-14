class AppState:
    """Classe para armazenar os dados globais do sistema."""
    
    def __init__(self):
        self.selected_process = None
        self.selected_cupom = None
        self.inscricao_estadual = None
        self.mes = None
        self.ano = None

    def set_data(self, selected_process, selected_cupom, inscricao_estadual, mes, ano):
        """Atualiza os dados no estado global"""
        self.selected_process = selected_process
        self.selected_cupom = selected_cupom
        self.inscricao_estadual = inscricao_estadual
        self.mes = mes
        self.ano = ano

# Criamos uma instância global da classe para armazenar os dados
app_state = AppState()
