import sys
import os
import subprocess
import tkinter as tk
from tkinter import simpledialog, messagebox
from playwright.sync_api import sync_playwright
import time

SISTEMAS = {
    1: {'nome': 'Digidoc', 'modulos': ['Distribuidores','Edoc','Comum','Documento','Protocolo', 'Externo','Pleno Administrativo','Revisores']},
    2: {'nome': 'Salus', 'modulos': ['Marcação de Consulta']},
    3: {'nome': 'SiaferjWeb', 'modulos': ['Serventia Judicial', 'Serventia Extrajudicial','Master','Magistrado','Judicial/Fotocópia','Fotocópia', 'Diretoria FERJ', 'Corregedoria','Controladoria', 'Administração de selos']},
    4: {'nome': 'Tutor', 'modulos': ['Alunos']},
    5: {'nome': 'Vínculos', 'modulos': ['Usuário da Vara', 'NPE', 'Laboratório de DNA']},
    6: {'nome': 'Accessus', 'modulos': ['Usuário de Setor', 'Supervisor', 'Secretário','Operacional', 'Magistrado']},
    7: {'nome': 'Haedus', 'modulos': ['Servidor TJMA']},
    8: {'nome': 'Jurisconsult', 'modulos': ['TJMA', 'TELETRABALHO', 'VEP_CNPJ', 'UNIDADE_MONITORAMENTO_CARCERARIO', 'SEJAP', 'SECRETÁRIOS JUDICIAIS-2G', 'SECRETÁRIOS JUDICIAIS - DISTRIBUIÇÃO',
        'SECRETÁRIOS JUDICIAIS', 'RELATÓRIO IBGE', 'Psicossocial', 'PLANEJAMENTO_ESTRATEGICO', 'PJE', 'OFICIAIS DE JUSTIÇA', 'MAGISTRADOS', 'JUIZADOS','Diretoria judiciária', 'DISTRIBUIÇÃO - 2° GRAU',
        'DIR_RH', 'DEFENSORIA/PROMOTORIA/OUTROS TRIBUNAIS','Cejusc','COORDENADORIA DA MULHER','CONTADORIA','CGJ', 'Avaliação Judicial','ATENDIMENTO', 'ASSESSORIA', 'ACESSOR_JUIZ_VEP']},
    9: {'nome': 'NexusRH', 'modulos': ['Servidores NexusRH','Magistrados NexusRH']},
    10: {'nome': 'ThemisSG Web', 'modulos': ['Consulta','Arquivo','Assessoria Jurídica da Presidência','Atendimento','Auditoria','Configuração',
        'Coordenação Câmaras','Coordenação Distribuição', 'Coordenação Processo Administrativo','Coordenação Protocolo', 'Distribuição','Estatística', 'Gabinete','Plantao Judiciario','Precatórios',
        'Protocolo', 'Protocolo Descentralizado','Secretaria Coordenação Câmaras','Secretaria da Distribuição', 'Virtualização']},
    11: {'nome': 'Termojuris', 'modulos': ['SECRETARIO', 'CEJUSC', 'CORREGEDORIA', 'MAGISTRADO',
            'MAGISTRADO_COGEX', 'SECRETARIO', 'SECRETARIO_COGEX']},
    12: {'nome': 'Convictus', 'modulos': ['Servidor VEP','Secretario VEP',
            'UMF', 'Núcleo de Suporte a Sistemas']},
    13: {'nome': 'Conciliação', 'modulos': ['Servidor Unidade']},
    14: {'nome': 'AR Digital', 'modulos': ['SERVIDOR_UNIDADE']},
    15: {'nome': 'Frottas', 'modulos': ['Solicitantes', 'Comum', 'Estatística', 'Operador Transporte', 'Transporte', 'Solicitantes']},
    16: {'nome': 'Gerenciador', 'modulos': ['SERVIÇOS', 'MODULOS DA CGJ', 'ATOS ADMINISTRATIVOS', 'INSTITUCIONAL', 'PLANTOES', 'NOTÍCIAS', 'PUBLICACOES', 'COMARCA VARAS TURMAS CAMARAS JUIZADOS SERVENTIAS']},
    17: {'nome': 'Auditus', 'modulos': ['Corregedoria', 'Usuário da Serventia', 'Secretário Judicial']},
    18: {'nome': 'Notarium', 'modulos': ['Corregedoria', 'Magistrados', 'Serventias extrajudiciais', 'Atendimento ao Público']},
    19: {'nome': 'CIG', 'modulos': ['Público', 'Suporte']},
    20: {'nome': 'ENGEDOC', 'modulos': ['Requisição - Cadastro', 'Fiscal', 'Requisição - Diretor', 'Requisição - Gestor']},
    21: {'nome': 'App Portal do Servidor', 'modulos': ['Servidores', 'Magistrados Aposentados', 'Delegatarios']},
    22: {'nome': 'IDS', 'modulos': ['Servidor']},
    23: {'nome': 'Materiales', 'modulos': ['Usuário de Setor']}
    
}

def verificar_playwright():
    try:
        from playwright.sync_api import sync_playwright
        return True
    except:
        return False

def instalar_playwright():
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "playwright"], check=True)
        subprocess.run([sys.executable, "-m", "playwright", "install"], check=True)
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao instalar Playwright: {str(e)}")
        return False

def configurar_ambientes():
    if getattr(sys, 'frozen', False):
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = os.path.join(sys._MEIPASS, 'playwright', 'browsers')

def obter_dados():
    root = tk.Tk()
    root.withdraw()
    sistemas_str = "\n".join([f"{k} - {v['nome']}" for k, v in SISTEMAS.items()])
    selecao = simpledialog.askstring("Sistemas", f"Quais sistemas deseja renovar? (Números separados por vírgula)\n\n{sistemas_str}", parent=root)

    if not selecao:
        messagebox.showerror("Erro", "Seleção de sistemas é obrigatória!")
        sys.exit(1)

    try:
        sistemas_selecionados = [int(s.strip()) for s in selecao.split(',')]
    except:
        messagebox.showerror("Erro", "Formato inválido! Use números separados por vírgula (ex: 1,3)")
        sys.exit(1)

    sistemas_para_renovar = []
    for num in sistemas_selecionados:
        if num in SISTEMAS:
            sistemas_para_renovar.append(SISTEMAS[num])
        else:
            messagebox.showerror("Erro", f"Sistema {num} não encontrado!")
            sys.exit(1)
    
    from datetime import datetime
    ano_futuro = datetime.now().year + 2
    data_expiracao = f"31/12/{ano_futuro}"

    dados = {
        'usuario': simpledialog.askstring("Entrada", "Digite seu usuário (CPF):", parent=root),
        'senha': simpledialog.askstring("Entrada", "Digite sua senha:", show='*', parent=root),
        'matricula': simpledialog.askstring("Entrada", "Digite a matrícula, CPF (sem pontos) ou nome completo do usuário:", parent=root),
        'data': data_expiracao,
        'autenticacao_codigo': simpledialog.askstring("Entrada", "Código de autenticação do sentinela (ou pressione Enter se não houver):", parent=root),
        'sistemas': sistemas_para_renovar,
        'usuario_glpi': simpledialog.askstring("Entrada", "Digite seu usuário do glpi", parent=root),
        'senha_glpi': simpledialog.askstring("Entrada", "Digite sua senha do glpi:", show='*', parent=root),
        'numero_solicitante': simpledialog.askstring("Entrada", "coloque o número do solicitante: ", parent=root)
    }

    if None in [dados['usuario'], dados['senha'], dados['matricula'], dados['data']]:
        messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
        sys.exit(1)

    return dados


print("Desenvolvido por Gustavo Marques\n")

class AutomacaoSentinela:
    def __init__(self, usuario: str, matricula: str, senha: str, data: str, autenticacao_codigo: str, sistemas: list, usuario_glpi: str, senha_glpi: str, numero_solicitante: int):
        self.usuario = usuario
        self.matricula = matricula
        self.senha = senha
        self.data = data
        self.autenticacao_codigo = autenticacao_codigo
        self.sistemas = sistemas
        self.usuario_glpi = usuario_glpi
        self.senha_glpi = senha_glpi
        self.numero_solicitante = numero_solicitante
        self.page = None
        self.sistemas_ok = []
        self.sistemas_falhos = []
        self.lotacao = ""
        self.nome_glpi = ""

    def executar(self):
        
        self.sistemas_falhos = []
        browser = p.chromium.launch(headless=False)
        self.page = browser.new_page()
        sistemas_nao_processados = []

        try:
            self.page.goto("https://sistemas.tjma.jus.br/sentinela/SistemaAction.welcome.mtw")
            print("Conectou ao sistema Sentinela")
            self.page.wait_for_load_state("domcontentloaded")
            self.inicio_tela()
            self.autenticador()
            self.procurar_usuario()
            self.capturar_nome()
            self.capturar_lotacao()
            self.renovar_status_usuario()

            for sistema in self.sistemas:
                print(f"\n=== Iniciando renovação para sistema: {sistema['nome']} ===")
                try:
                    self.page.click("xpath=//a[contains(@href, 'SistemaAction.welcome.mtw') or contains(text(), 'Sentinela')]", timeout=5000)
                    time.sleep(1)
                    self.procurar_usuario()
                    self.renovar_sistema(sistema)

                    print(f"\n=== Processo concluído para sistema: {sistema['nome']} ===\n")
                except Exception as e:
                    print(f"Erro ao processar sistema {sistema['nome']}: {e}")
                    sistemas_nao_processados.append(sistema['nome'])
                    continue
            for nome in self.sistemas_falhos:
                if nome not in sistemas_nao_processados:
                    sistemas_nao_processados.append(nome)

            if not sistemas_nao_processados:
                print("Todas as renovações foram concluídas com sucesso!")
            else:
                print("\nAlguns sistemas não puderam ser renovados:")
                for nome in sistemas_nao_processados:
                    print(f"- {nome}")

        except Exception as e:
            print(f"Erro durante a execução: {e}")
            raise

    def inicio_tela(self):
        try:
            print("Logando no sistema...")
            self.page.fill("xpath=//input[@id='usuario']", self.usuario)
            self.page.fill("xpath=//input[@id='senha']", self.senha)
            self.page.click("xpath=//input[@name='entrar']")
            print("Login realizado com sucesso")
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao logar: {e}")
            raise

    def autenticador(self):
        try:
            print("Verificando se há necessidade de autenticação...")
            self.page.wait_for_selector("xpath=//*[@id='codigo']", timeout=3000)

            if self.autenticacao_codigo.strip():
                print("Preenchendo código de autenticação...")
                self.page.fill("xpath=//*[@id='codigo']", self.autenticacao_codigo)
                self.page.click("xpath=//*[@id='campos']/form/div[5]/input")
                print("Autenticação concluída. Continuando...")
                time.sleep(1)
            else:
                print("Código de autenticação não fornecido. Pulando etapa.")
        except:
            print("Campo de autenticação não encontrado. Continuando...")

    def procurar_usuario(self):
        try:
            print("\nProcurando usuário...")
            self.page.hover("xpath=//a[normalize-space()='cadastros']")
            self.page.hover("xpath=//a[@href='/sentinela/UsuarioAction.preConsultar.mtw']")
            self.page.click("xpath=//a[@href='/sentinela/UsuarioAction.preConsultar.mtw']")
            print("Acessou a consulta de usuários")
            time.sleep(1)

            identificador = self.matricula.strip()
            if identificador.isdigit():
                if len(identificador) == 11:
                    self.page.fill("xpath=//input[@id='strUsuarioId']", identificador)
                    print(f"Busca por CPF: {identificador}")
                else:
                    self.page.fill("xpath=//input[@id='intMatricula']", identificador)
                    print(f"Busca por matrícula: {identificador}")
            else:
                self.page.fill("xpath=//input[@id='strUsuario']", identificador)
                print(f"Busca por nome: {identificador}")
            self.page.click('xpath=//input[@name="btnOk"]')
            print("Clicou no botão OK")
            time.sleep(1)

            print("Editando usuário...")
            self.page.click("xpath=//img[@title='Editar Usuário']")
            print("Clicou no ícone de edição")
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao preencher ou interagir: {e}")
            raise

    def capturar_nome(self):
        try:
            self.nome_glpi = self.page.input_value("xpath=//input[@id='strEmailUsuarioAltera']")
            print(f"nome capturado: {self.nome_glpi}")
        except:
            print("Não foi possível capturar o nome")

    def capturar_lotacao(self):
        try:
            self.lotacao = self.page.input_value("xpath=//input[@id='strLotacaoSentinela']")
            print(f"Lotação capturada: {self.lotacao}")
        except:
            print("Não foi possível capturar a lotação")

    def renovar_status_usuario(self):
        try:
            print("\nRenovando status do usuário...")
            self.page.click("xpath=//*[@id='indStatusUsuario']")
            print("Confirmando alteração do status")
            frame = self.page.frame_locator("iframe").first
            frame.locator("body#tinymce").click()
            self.page.keyboard.type("Alteração.")
            self.page.click("xpath=//a[@id='link_submit']")
            time.sleep(1)
            self.page.click("xpath=//div[@class='tmask']")
            print("Alteração confirmada com sucesso")
            time.sleep(0.5)
            self.page.click("xpath=//img[@title='Editar Usuário']")
        except Exception as e:
            print(f"Erro ao renovar status do usuário: {e}")
            raise

    def renovar_sistema(self, sistema):
        try:
            nome_sistema = sistema['nome']
            modulos_esperados = sistema['modulos']
            print(f"\nAcessando sistema: {nome_sistema}")

            try:
                self.page.click(f"xpath=//td[normalize-space()='{nome_sistema}']", timeout=5000)
                print(f"Acessou a aba {nome_sistema}")
            except Exception as e:
                print(f"Erro ao acessar sistema {nome_sistema}: {e}")
                self.sistemas_falhos.append(nome_sistema)
                return

            time.sleep(1)
            try:
                self.page.wait_for_selector(f"xpath=//td[contains(text(), '{nome_sistema}')]", timeout=5000)
            except:
                print(f"Não conseguiu confirmar acesso ao sistema {nome_sistema}")
                self.page.screenshot(path=f"erro_confirmacao_{nome_sistema}.png")
                self.sistemas_falhos.append(nome_sistema)
                return

            modulos_encontrados = []
            for modulo in modulos_esperados:
                try:
                    if self.page.is_visible(f"xpath=//tr[td[normalize-space()='{modulo}']]", timeout=2000):
                        modulos_encontrados.append(modulo)
                        print(f"Encontrado módulo: {modulo}")
                except:
                    continue

            if not modulos_encontrados:
                print("Nenhum módulo encontrado para este sistema.")
                self.sistemas_falhos.append(nome_sistema)
                return

            print(f"\nIniciando renovação para {len(modulos_encontrados)} módulos...")

            for modulo in modulos_encontrados:
                try:
                    print(f"\nProcessando módulo: {modulo}")
                    linhas = self.page.query_selector_all(f"xpath=//tr[td[normalize-space()='{modulo}']]")
                    linha_valida = None

                    for linha in linhas:
                        if linha.is_visible():
                            linha_valida = linha
                            break

                    if not linha_valida:
                        print(f"Linha do módulo {modulo} não encontrada")
                        continue

                    lapis = None
                    tentativas = [
                        ".//img[contains(@src, 'editar') or contains(@alt, 'editar')]",
                        ".//a[contains(@href, 'editar')]",
                        ".//img[contains(@title, 'Editar')]",
                        ".//a[img[contains(@src, 'editar')]]",
                        ".//*[contains(@class, 'editar')]",
                        ".//img[@title='Editar Vínculo']"
                    ]

                    for tentativa in tentativas:
                        lapis = linha_valida.query_selector(f"xpath={tentativa}")
                        if lapis:
                            break

                    if not lapis:
                        print(f"Ícone de edição não encontrado para {modulo}")
                        self.page.screenshot(path=f"erro_lapis_{nome_sistema}_{modulo.replace('/', '_')}.png")
                        continue

                    try:
                        lapis.scroll_into_view_if_needed()
                        lapis.click()
                        print(f"Clicou no ícone de edição para {modulo}")
                        time.sleep(1)

                        self.page.fill("xpath=//input[@id='dtaExpiracaoGrupo']", self.data)
                        print(f"Data de validade alterada para {self.data}")

                        self.page.click("xpath=//span[@class='checked']//input[@id='indUsuarioAtivoGrupo']")
                        if nome_sistema == "App Portal do Servidor":
                            self.page.fill("xpath=//input[@id='strDocumentoOficial']", "1")
                            print("Preencheu o documento oficial com 1")
                        self.page.click("xpath=//a[@id='link_submit_inferior']")
                        print("Alterações salvas")
                        time.sleep(1)

                        try:
                            self.page.click("xpath=//div[@class='tmask']", timeout=2000)
                            print("Alerta fechado")
                            self.page.click("xpath=//img[@title='Editar Vínculo']")
                            print("Voltou para a tela do sistema")
                            time.sleep(1)
                        except:
                            pass

                        self.page.click(f"xpath=//td[normalize-space()='{nome_sistema}']")
                        print(f"Voltou para a tela do {nome_sistema}")
                        time.sleep(1)

                    except Exception as e:
                        print(f"Erro ao editar módulo {modulo}: {e}")
                        self.page.screenshot(path=f"erro_edicao_{nome_sistema}_{modulo.replace('/', '_')}.png")
                        continue

                except Exception as e:
                    print(f"Erro ao processar módulo {modulo}: {e}")
                    self.page.screenshot(path=f"erro_{modulo.replace('/', '_')}.png")
                    continue

            print("\nProcesso de renovação concluído para todos os módulos encontrados!")
            self.sistemas_ok.append(nome_sistema)

        except Exception as e:
            print(f"Erro geral no processamento do {nome_sistema}: {e}")

    def login_glpi(self):
        try:
            self.page.goto("https://atendimento.tjma.jus.br")
            print("logando no glpi")
            self.page.fill("xpath=//input[@id='login_name']", self.usuario_glpi)
            self.page.fill("xpath=//input[@id='login_password']", self.senha_glpi)
            self.page.click("xpath=//button[@name='submit']")
            print("logado com sucesso")
            time.sleep(2)

        except Exception as e:
            print(f"Erro ao logar: {e}")
            raise

    def criar_formulario(self):
        try:
            print('criando um formulario no glpi')
            if self.page.is_visible("xpath=//span[normalize-space()='Formulários']"):
                print("Formulários já visível, clicando direto...")
                self.page.click("xpath=//span[normalize-space()='Formulários']")
            else:
                print("Formulários não visível, abrindo menu headset...")
                self.page.hover("xpath=//i[@class='fa-fw ti ti-headset']")
                self.page.click("xpath=//i[@class='fa-fw ti ti-headset']")
                self.page.wait_for_selector("xpath=//span[normalize-space()='Formulários']", timeout=5000)
                self.page.click("xpath=//span[normalize-space()='Formulários']")
            
            time.sleep(1)
            self.page.click("xpath=//a[normalize-space()='Registre sua Requisição (Formulário do Técnico)']")
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao criar o formulário: {e}")
            raise   

    def criacao_requisicao(self, nome_sistema):
        try:
            print('criando requisição no glpi')
            titulo = f"Renovação de acesso ao sistema {nome_sistema}"
            self.page.click("xpath=//h1[@class='form-title']")
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            self.page.keyboard.type(self.nome_glpi)
            time.sleep(7)
            self.page.keyboard.press("Enter")
            print('colocou o nome da requerente')
            time.sleep(0.5)
            self.page.click("xpath=//label[@title='Sim']")
            print('aplicou a tarefa para si mesmo')
            time.sleep(0.5)
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            self.page.keyboard.type(self.numero_solicitante)
            print('preencheu o numero da solicitante')
            time.sleep(0.5)
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            self.page.keyboard.type(self.lotacao)
            time.sleep(5)
            self.page.keyboard.press("Enter")
            print("colocou a lotacao da solicitante")
            time.sleep(0.5)
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            sistemas_sem_glpi = ["App Portal do Servidor", "IDS", "ThemisSG Web", "Accessus", "CIG", "Planus" ]
            if nome_sistema in sistemas_sem_glpi:
                self.page.keyboard.type("Sentinela > acesso")
                print(f"colocou Sentinela > alteração (exceção para {nome_sistema})")
            else:
                self.page.keyboard.type(f"{nome_sistema} > acesso")
                print(f"colocou o sistema {nome_sistema}")
            time.sleep(4)
            self.page.keyboard.press("Enter")
            print(f"colocou o sistema {nome_sistema}")
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            self.page.keyboard.type(titulo)
            print("colocou o titulo")
            time.sleep(0.5)
            self.page.keyboard.press("Tab") 
            self.page.keyboard.press("Enter")
            self.page.keyboard.type(titulo)
            print('Preencheu a descrição')
            time.sleep(1)
            self.page.click('xpath=//*[@id="plugin_formcreator_form"]/div[1]/button')
            time.sleep(1)
            
        except Exception as e:
            print(f"Erro ao criar a requisição: {e}")
            raise 

    def fechamento_requisicao(self,nome_sistema):
        try:
            solucao = f"Foi feito a renovação de validade do sistema {nome_sistema}."
            print("fechando os chamados em lote")
            self.page.click("xpath=//a[normalize-space()='Home']")
            time.sleep(1)
            print("clicou em home")
            self.page.click("xpath=//a[@title='Visão pessoal']")
            time.sleep(1)
            print("clicou em visão pessoal")
            self.page.click(f"xpath=//span[contains(text(),'Renovação de acesso ao sistema {nome_sistema}')]")
            time.sleep(0.5)
            print("clicando no chamado")
            self.page.click("xpath=//span[normalize-space()='Solução']")
            time.sleep(0.5)
            print("clicou em solução")
            self.page.click('xpath=//*[@id="new-ITILSolution-block"]/div/div[2]/div/div/div[2]/div/form/div[1]/div[1]/div/div[1]/div[1]/div[2]')
            self.page.wait_for_load_state('load', timeout=5000)
            time.sleep(0.5)
            self.page.keyboard.type(solucao)
            print("digitou a solução")
            self.page.click('xpath=//*[@id="new-ITILSolution-block"]/div/div[2]/div/div/div[2]/div/form/div[2]/button')
            time.sleep(1)

        except Exception as e:
            print(f"Erro ao fechar a requisição: {e}")
            raise 

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    try:
        configurar_ambientes()
        
        if not verificar_playwright():
            if not messagebox.askyesno("Instalação", "Playwright não encontrado. Deseja instalar agora?"):
                sys.exit(1)
            if not instalar_playwright():
                sys.exit(1)
        
        tk.messagebox.showinfo("Sucesso", "Automação desenvolvida por Gustavo Marques!")
        dados = obter_dados()
        bot = AutomacaoSentinela(**dados)
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            bot.executar()
            if bot.sistemas_ok:
                bot.page = browser.new_page()
                bot.login_glpi()
                for sistema in bot.sistemas_ok:
                    bot.criar_formulario()
                    bot.criacao_requisicao(sistema)
                    bot.fechamento_requisicao(sistema)

        if not bot.sistemas_falhos:
            tk.messagebox.showinfo("Sucesso", "Todas as renovações foram concluídas com sucesso!")
        else:
            falhas = "\n".join(f"- {sistema}" for sistema in bot.sistemas_falhos)
            tk.messagebox.showwarning("Aviso", f"Alguns sistemas não puderam ser renovados:\n\n{falhas}")

    except Exception as e:
        tk.messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    finally:
        root.destroy()
