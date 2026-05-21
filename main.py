import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from pathlib import Path
import asyncio
import re
import subprocess
import os
from datetime import datetime
from playwright.async_api import async_playwright
from openpyxl import load_workbook
from docx import Document
from config import (
    SEI_URL, LOCATOR_USUARIO, LOCATOR_SENHA, LOCATOR_ORGAO,
    LOCATOR_BOTAO_ACESSAR, LOCATOR_PESQUISA_RAPIDA, LOCATOR_BOTAO_PESQUISA,
    LOCATOR_FECHAR_JANELA, ORGAO_PADRAO, COLUNA_ARQUIVO_DOCX, COLUNA_ARQUIVO_PDF,
    HEADLESS, TIMEOUT, COLUNA_PESQUISA, LINHA_PESQUISA
)
def obter_locator(page, config_locator):
    """
    Obtém um locator da página baseado na configuração
    
    Args:
        page: Página do Playwright
        config_locator: Dicionário com tipo e parâmetros do locator
    
    Returns:
        Locator do Playwright
    """
    if config_locator['tipo'] == 'role':
        return page.get_by_role(config_locator['role'], name=config_locator['nome'])
    elif config_locator['tipo'] == 'selector':
        return page.locator(config_locator['selector'])
    else:
        raise ValueError(f"Tipo de locator não suportado: {config_locator['tipo']}")


def copiar_para_clipboard(texto):
    """
    Copia texto para o clipboard do Windows
    
    Args:
        texto: Texto a ser copiado
    """
    try:
        # Usa set-clipboard do PowerShell
        cmd = f'@"\n{texto}\n"@ | Set-Clipboard'
        subprocess.run(['powershell', '-Command', cmd], check=True, capture_output=True)
    except Exception as e:
        print(f"Erro ao copiar para clipboard: {e}")


def extrair_numero_relatorio(caminho_arquivo):
    """
    Extrai o número do relatório técnico a partir do caminho do arquivo.
    
    Exemplos:
    - "...gfe_rf_2026_27_fisc_fmsb_alto_rio_doce.docx" -> "27 GFE – ALTO RIO DOCE"
    - "...gfe_rf_2026_28_fisc_fmsb_senador_modestino_goncalves.docx" -> "28 GFE – SENADOR MODESTINO GONCALVES"
    
    Args:
        caminho_arquivo: Caminho completo do arquivo .docx
    
    Returns:
        String formatada "XX GFE – NOME"
    """
    try:
        nome_arquivo = Path(caminho_arquivo).stem  # Remove .docx
        
        # Padrão: gfe_rf_2026_XX_fisc_fmsb_descrição
        match = re.search(r'gfe_rf_\d+_(\d+)_fisc_fmsb_(.+)', nome_arquivo)
        if match:
            numero = match.group(1)
            descricao = match.group(2)
            
            # Formata a descrição (substitui _ por espaço e coloca em maiúsculas)
            descricao_formatada = descricao.replace('_', ' ').upper()
            
            return f"{numero} GFE – {descricao_formatada}"
    except Exception as e:
        print(f"Erro ao extrair número do relatório: {e}")
    
    return ""


def limpar_documento_word(caminho_arquivo):
    """
    Extrai todo o conteúdo do documento Word a partir da página 2.
    Preserva a formatação (negrito, itálico, cores, etc.) usando HTML.
    
    Args:
        caminho_arquivo: Caminho do arquivo .docx
    
    Returns:
        Conteúdo em HTML preservando a formatação
    """
    try:
        doc = Document(caminho_arquivo)
        
        # Extrai parágrafos mantendo formatação
        html_content = []
        first_para = True
        
        for para in doc.paragraphs:
            # Pula o primeiro parágrafo (primeira página)
            if first_para:
                first_para = False
                continue
            
            # Processa o parágrafo mantendo formatação dos runs
            if para.text.strip():  # Apenas parágrafos com conteúdo
                para_html = "<p>"
                
                for run in para.runs:
                    if run.text:
                        texto = run.text.replace("<", "&lt;").replace(">", "&gt;")
                        
                        # Aplica formatação (negrito, itálico, sublinhado)
                        if run.bold:
                            texto = f"<strong>{texto}</strong>"
                        if run.italic:
                            texto = f"<em>{texto}</em>"
                        if run.underline:
                            texto = f"<u>{texto}</u>"
                        
                        para_html += texto
                
                para_html += "</p>"
                html_content.append(para_html)
            else:
                # Preserva parágrafos vazios
                html_content.append("<br>")
        
        # Retorna o conteúdo em HTML
        return "\n".join(html_content)
    except Exception as e:
        print(f"Erro ao processar documento Word: {e}")
        return ""


def obter_data_formatada():
    """Retorna a data atual formatada (DD/MM/YYYY)"""
    return datetime.now().strftime("%d/%m/%Y")


class AutomacaoSEI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação SEI - ARSAE-MG")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # Variáveis
        self.usuario = tk.StringVar()
        self.senha = tk.StringVar()
        self.caminho_arquivo = tk.StringVar()
        
        self.criar_interface()
    
    def criar_interface(self):
        """Cria a interface Tkinter"""
        # Frame principal
        frame_principal = tk.Frame(self.root, padx=20, pady=20)
        frame_principal.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = tk.Label(frame_principal, text="Dados de Acesso SEI", font=("Arial", 14, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Campo Usuário
        tk.Label(frame_principal, text="Usuário:", font=("Arial", 10)).pack(anchor=tk.W)
        entry_usuario = tk.Entry(frame_principal, textvariable=self.usuario, width=40)
        entry_usuario.pack(pady=(0, 10), fill=tk.X)
        
        # Campo Senha
        tk.Label(frame_principal, text="Senha:", font=("Arial", 10)).pack(anchor=tk.W)
        entry_senha = tk.Entry(frame_principal, textvariable=self.senha, width=40, show="•")
        entry_senha.pack(pady=(0, 10), fill=tk.X)
        
        # Campo Arquivo
        tk.Label(frame_principal, text="Arquivo .xlsx:", font=("Arial", 10)).pack(anchor=tk.W)
        frame_arquivo = tk.Frame(frame_principal)
        frame_arquivo.pack(pady=(0, 20), fill=tk.X)
        
        entry_arquivo = tk.Entry(frame_arquivo, textvariable=self.caminho_arquivo, width=28)
        entry_arquivo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_browse = tk.Button(frame_arquivo, text="Procurar", command=self.procurar_arquivo)
        btn_browse.pack(side=tk.LEFT, padx=(10, 0))
        
        # Botão Iniciar
        btn_iniciar = tk.Button(frame_principal, text="INICIAR AUTOMAÇÃO", 
                               command=self.iniciar_automacao, 
                               bg="#4CAF50", fg="white", 
                               font=("Arial", 11, "bold"),
                               height=2)
        btn_iniciar.pack(fill=tk.X, pady=10)
        
        # Botão Sair
        btn_sair = tk.Button(frame_principal, text="Sair", 
                            command=self.root.quit,
                            bg="#f44336", fg="white",
                            font=("Arial", 10))
        btn_sair.pack(fill=tk.X)
    
    def procurar_arquivo(self):
        """Abre o diálogo para procurar arquivo .xlsx"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo .xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if arquivo:
            self.caminho_arquivo.set(arquivo)
    
    def validar_entrada(self):
        """Valida os campos de entrada"""
        if not self.usuario.get().strip():
            messagebox.showerror("Erro", "Usuário não pode estar vazio!")
            return False
        
        if not self.senha.get():
            messagebox.showerror("Erro", "Senha não pode estar vazia!")
            return False
        
        if not self.caminho_arquivo.get().strip():
            messagebox.showerror("Erro", "Caminho do arquivo não pode estar vazio!")
            return False
        
        if not Path(self.caminho_arquivo.get()).exists():
            messagebox.showerror("Erro", "Arquivo não encontrado!")
            return False
        
        return True
    
    def iniciar_automacao(self):
        """Inicia a automação em uma thread separada"""
        if not self.validar_entrada():
            return
        
        # Desabilita botão enquanto executa
        for widget in self.root.winfo_children():
            for child in widget.winfo_children():
                if isinstance(child, tk.Button):
                    child.config(state=tk.DISABLED)
        
        # Executa em thread separada
        thread = threading.Thread(target=self.executar_automacao)
        thread.daemon = True
        thread.start()
    
    def executar_automacao(self):
        """Executa a automação com Playwright"""
        try:
            # Executa o loop assíncrono
            asyncio.run(self.automacao_sei())
            messagebox.showinfo("Sucesso", "Automação concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a automação:\n{str(e)}")
        finally:
            # Reabilita botões
            for widget in self.root.winfo_children():
                for child in widget.winfo_children():
                    if isinstance(child, tk.Button):
                        child.config(state=tk.NORMAL)
    
    async def automacao_sei(self):
        """Automação do site SEI com Playwright"""
        async with async_playwright() as p:
            # Abre navegador
            print("Abrindo navegador Chrome...")
            browser = await p.chromium.launch(headless=HEADLESS)
            page = await browser.new_page()
            
            try:
                # 1. Acessa site SEI
                print(f"Acessando {SEI_URL}...")
                await page.goto(SEI_URL, wait_until="domcontentloaded", timeout=TIMEOUT)
                print("Página carregada, aguardando elementos...")
                await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                await page.wait_for_timeout(2000)  # Aguarda 2s extras para elementos carregar
                
                # 2. Preenche usuário
                print("Verificando campo de usuário...")
                locator_usuario = obter_locator(page, LOCATOR_USUARIO)
                try:
                    await locator_usuario.wait_for(timeout=10000)
                    print("✓ Campo de usuário encontrado")
                    print("Preenchendo usuário...")
                    await locator_usuario.fill(self.usuario.get())
                    await page.wait_for_timeout(1000)
                except Exception as e:
                    print(f"✗ Erro no campo de usuário: {e}")
                    await page.screenshot(path="erro_usuario_nao_encontrado.png")
                    raise Exception("Campo de usuário não encontrado. Verifique erro_usuario_nao_encontrado.png")
                
                # 3. Preenche senha
                print("Verificando campo de senha...")
                locator_senha = obter_locator(page, LOCATOR_SENHA)
                try:
                    await locator_senha.wait_for(timeout=10000)
                    print("✓ Campo de senha encontrado")
                    print("Preenchendo senha...")
                    await locator_senha.fill(self.senha.get())
                    await page.wait_for_timeout(1000)
                except Exception as e:
                    print(f"✗ Erro no campo de senha: {e}")
                    await page.screenshot(path="erro_senha_nao_encontrada.png")
                    raise Exception("Campo de senha não encontrado. Verifique erro_senha_nao_encontrada.png")
                
                # 4. Seleciona órgão
                print(f"Verificando campo de órgão...")
                locator_orgao = obter_locator(page, LOCATOR_ORGAO)
                try:
                    await locator_orgao.wait_for(timeout=10000)
                    print(f"Selecionando {ORGAO_PADRAO}...")
                    await locator_orgao.select_option(label=ORGAO_PADRAO)
                    print(f"✓ {ORGAO_PADRAO} selecionado")
                except Exception as e:
                    print(f"Erro ao selecionar órgão: {e}")
                    await page.screenshot(path="erro_orgao.png")
                    raise Exception(f"Não foi possível selecionar {ORGAO_PADRAO}. Verifique erro_orgao.png")
                
                await page.wait_for_timeout(500)
                
                # 5. Clica em ACESSAR
                print("Localizando botão ACESSAR...")
                locator_acessar = obter_locator(page, LOCATOR_BOTAO_ACESSAR)
                try:
                    await locator_acessar.wait_for(timeout=10000)
                    print("Clicando em ACESSAR...")
                    await locator_acessar.click()
                    print("Aguardando carregamento da página após login...")
                    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                    await page.wait_for_timeout(3000)  # Aguarda 3s após login
                    
                    # Clica no link "Fechar janela (ESC)"
                    print("Fechando janela de boas-vindas...")
                    locator_fechar = obter_locator(page, LOCATOR_FECHAR_JANELA)
                    try:
                        await locator_fechar.wait_for(timeout=5000)
                        await locator_fechar.click()
                        await page.wait_for_timeout(500)
                    except:
                        print("Link de fechar janela não encontrado, continuando...")
                    
                except Exception as e:
                    print(f"Erro ao clicar em ACESSAR: {e}")
                    await page.screenshot(path="erro_botao_acessar.png")
                    raise Exception("Não foi possível clicar no botão ACESSAR. Verifique erro_botao_acessar.png")
                
                # 6. Lê o arquivo .xlsx
                print("Lendo arquivo Excel...")
                caminho = self.caminho_arquivo.get()
                workbook = load_workbook(caminho)
                worksheet = workbook.active
                
                # 7. Copia valor da célula
                celula = f"{COLUNA_PESQUISA}{LINHA_PESQUISA}"
                valor_pesquisa = worksheet[celula].value
                
                if valor_pesquisa is None:
                    raise ValueError(f"Célula {celula} está vazia!")
                
                print(f"Valor encontrado para pesquisa: {valor_pesquisa}")
                
                # ============================================================
                # LÊ DADOS DO ARQUIVO EXCEL - Colunas E e F
                # ============================================================
                print("Obtendo caminhos de arquivos do Excel...")
                
                # Coluna E - Caminho do arquivo .docx
                caminho_docx = worksheet[f"{COLUNA_ARQUIVO_DOCX}{LINHA_PESQUISA}"].value
                if not caminho_docx:
                    raise ValueError(f"Célula E{LINHA_PESQUISA} (arquivo .docx) está vazia!")
                
                # Coluna F - Caminho do arquivo .pdf
                caminho_pdf = worksheet[f"{COLUNA_ARQUIVO_PDF}{LINHA_PESQUISA}"].value
                if not caminho_pdf:
                    raise ValueError(f"Célula F{LINHA_PESQUISA} (arquivo .pdf) está vazia!")
                
                print(f"Arquivos encontrados:")
                print(f"  - Documento DOCX: {caminho_docx}")
                print(f"  - Documento PDF: {caminho_pdf}")
                
                # ============================================================
                # ETAPA 1: PESQUISA RÁPIDA
                # ============================================================
                print("\n[ETAPA 1] Executando pesquisa rápida...")
                locator_pesquisa = obter_locator(page, LOCATOR_PESQUISA_RAPIDA)
                try:
                    await locator_pesquisa.wait_for(timeout=10000)
                    print("Preenchendo campo de pesquisa...")
                    await locator_pesquisa.fill(str(valor_pesquisa))
                    await page.wait_for_timeout(500)
                    
                    print("Clicando no botão de pesquisa...")
                    locator_botao_pesquisa = obter_locator(page, LOCATOR_BOTAO_PESQUISA)
                    await locator_botao_pesquisa.wait_for(timeout=5000)
                    await locator_botao_pesquisa.click()
                    
                    print("Aguardando resultados...")
                    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                except Exception as e:
                    print(f"Erro ao pesquisar: {e}")
                    await page.screenshot(path="erro_pesquisa.png")
                    raise Exception("Não foi possível realizar a pesquisa. Verifique erro_pesquisa.png")
                
                # ============================================================
                # ETAPA 2: CLICAR NO DOCUMENTO NA ÁRVORE
                # ============================================================
                print("\n[ETAPA 2] Clicando no documento na árvore...")
                try:
                    # Extrai o número do relatório do caminho do arquivo
                    numero_relatorio = extrair_numero_relatorio(caminho_docx)
                    print(f"  - Número extraído: {numero_relatorio}")
                    
                    # Constrói o nome do link dinâmico
                    # Exemplo: "Relatório Técnico 28 GFE ("
                    nome_link = f"Relatório Técnico {numero_relatorio.split()[0]} GFE ("
                    print(f"  - Procurando link: '{nome_link}'")
                    
                    # Acessa o iframe da árvore e clica no link
                    iframe_arvore = page.locator('iframe[name="ifrArvore"]')
                    await iframe_arvore.wait_for(timeout=10000)
                    frame_arvore = iframe_arvore.content_frame
                    
                    # Encontra o link com o nome dinâmico
                    link_documento = frame_arvore.get_by_role("link").filter(has_text=nome_link)
                    await link_documento.wait_for(timeout=5000)
                    print(f"  - Link encontrado, clicando...")
                    await link_documento.click()
                    
                    print("✓ Documento aberto")
                    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                    await page.wait_for_timeout(2000)
                    
                except Exception as e:
                    print(f"✗ Erro ao clicar no documento: {e}")
                    await page.screenshot(path="erro_clicar_documento_arvore.png")
                    raise Exception(f"Não foi possível clicar no documento. Verifique erro_clicar_documento_arvore.png")
                
                # ============================================================
                # ETAPA 3: EDITAR CONTEÚDO (ABRE NOVA JANELA)
                # ============================================================
                print("\n[ETAPA 3] Editando conteúdo do documento...")
                
                # Acessa iframe do conteúdo e clica em "Editar Conteúdo"
                try:
                    iframe_conteudo = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    await iframe_conteudo.wait_for(timeout=10000)
                    frame_conteudo = iframe_conteudo.content_frame
                    
                    link_editar = frame_conteudo.get_by_role("link", name="Editar Conteúdo")
                    await link_editar.wait_for(timeout=5000)
                    
                    # Prepara para detectar nova página/popup
                    print("  - Clicando em 'Editar Conteúdo'...")
                    async with page.context.expect_page() as new_page_event:
                        await link_editar.click()
                    
                    # Obtém a nova página aberta
                    nova_pagina = await new_page_event.value
                    print("✓ Nova janela de edição aberta")
                    
                    # Aguarda o carregamento da nova janela
                    await nova_pagina.wait_for_load_state("domcontentloaded", timeout=TIMEOUT)
                    await nova_pagina.wait_for_timeout(2000)
                    
                    # Atualiza a referência 'page' para a nova janela
                    page = nova_pagina
                    print("✓ Contexto alterado para a nova janela")
                    
                except Exception as e:
                    print(f"✗ Erro ao clicar em 'Editar Conteúdo': {e}")
                    await page.screenshot(path="erro_editar_conteudo.png")
                    raise Exception("Não foi possível clicar em 'Editar Conteúdo'")
                
                # ============================================================
                # ETAPA 4: LIMPAR CONTEÚDO ANTERIOR
                # ============================================================
                print("\n[ETAPA 4] Limpando conteúdo anterior...")
                try:
                    iframe_corpo = page.locator('iframe[title="Corpo do Texto"]')
                    await iframe_corpo.wait_for(timeout=10000)
                    frame_corpo = iframe_corpo.content_frame
                    
                    # Seleciona e deleta o texto padrão
                    texto_padrao = frame_corpo.get_by_text("[Inserir texto]")
                    await texto_padrao.wait_for(timeout=5000)
                    
                    # Seleciona tudo (Ctrl+A) e deleta
                    await frame_corpo.locator('body').click()
                    await page.keyboard.press("Control+A")
                    await page.keyboard.press("Delete")
                    print("✓ Conteúdo anterior removido")
                    
                except Exception as e:
                    print(f"⚠️  Aviso ao limpar conteúdo: {e}")
                    # Continua mesmo se não conseguir limpar
                
                # ============================================================
                # ETAPA 4 (revisado): LIMPAR E PREPARAR CONTEÚDO
                # ============================================================
                print("\n[ETAPA 4] Preparando para inserir novo conteúdo...")
                try:
                    # Seleciona todo o conteúdo no iframe
                    iframe_corpo = page.locator('iframe[title="Corpo do Texto"]')
                    await iframe_corpo.wait_for(timeout=10000)
                    frame_corpo = iframe_corpo.content_frame
                    
                    if frame_corpo is None:
                        raise Exception("Não foi possível acessar o frame do corpo do texto")
                    
                    # Clica no body
                    body_element = frame_corpo.locator('body')
                    await body_element.click()
                    await page.wait_for_timeout(300)
                    
                    # Seleciona tudo (Ctrl+A)
                    await page.keyboard.press("Control+A")
                    await page.wait_for_timeout(200)
                    
                    print("✓ Conteúdo anterior selecionado")
                    
                except Exception as e:
                    print(f"⚠️  Aviso ao preparar: {e}")
                    # Continua mesmo se houver erro
                
                # ============================================================
                # ETAPA 5: ABRIR WORD E COPIAR CONTEÚDO
                # ============================================================
                print("\n[ETAPA 5] Abrindo documento Word e copiando conteúdo...")
                try:
                    # Abre o documento Word com aplicação padrão (Windows)
                    print(f"  - Abrindo {caminho_docx}...")
                    os.startfile(caminho_docx)
                    await page.wait_for_timeout(5000)  # Aguarda abrir e carregar
                    
                    # Executa PowerShell para ir até a página 2 e copiar
                    print("  - Navegando até página 2 com Page Down...")
                    subprocess.run([
                        'powershell',
                        '-Command',
                        '''
                        Add-Type -AssemblyName System.Windows.Forms
                        
                        # Aguarda um pouco após abrir
                        Start-Sleep -Milliseconds 1000
                        
                        # Pressiona Page Down primeira vez
                        [System.Windows.Forms.SendKeys]::SendWait("{Page_Down}")
                        Start-Sleep -Milliseconds 800
                        
                        # Pressiona Page Down segunda vez
                        [System.Windows.Forms.SendKeys]::SendWait("{Page_Down}")
                        Start-Sleep -Milliseconds 1000
                        
                        # Aguarda antes de selecionar
                        Start-Sleep -Milliseconds 500
                        
                        # Seleciona do cursor até o fim (Ctrl+Shift+End)
                        [System.Windows.Forms.SendKeys]::SendWait("^+{End}")
                        Start-Sleep -Milliseconds 800
                        
                        # Copia (Ctrl+C)
                        [System.Windows.Forms.SendKeys]::SendWait("^c")
                        Start-Sleep -Milliseconds 800
                        '''
                    ], check=True, capture_output=True)
                    
                    print("✓ Conteúdo de página 2 até o fim foi copiado")
                    await page.wait_for_timeout(1000)
                    
                except Exception as e:
                    print(f"✗ Erro ao abrir Word: {e}")
                    raise Exception(f"Não foi possível processar {caminho_docx}")
                
                # ============================================================
                # ETAPA 6: COLAR CONTEÚDO NA JANELA DE EDIÇÃO
                # ============================================================
                print("\n[ETAPA 6] Colando conteúdo do clipboard...")
                try:
                    # Localiza o iframe do corpo do texto
                    print("  - Localizando caixa de edição...")
                    iframe_corpo = page.locator('iframe[title="Corpo do Texto"]')
                    await iframe_corpo.wait_for(timeout=10000)
                    frame_corpo = iframe_corpo.content_frame
                    
                    if frame_corpo is None:
                        raise Exception("Não foi possível acessar o frame do corpo do texto")
                    
                    # Seleciona a caixa de texto
                    print("  - Selecionando caixa de edição...")
                    body_element = frame_corpo.locator('body')
                    await body_element.click()
                    await page.wait_for_timeout(500)
                    
                    # Seleciona todo o conteúdo (Ctrl+A)
                    print("  - Selecionando conteúdo anterior...")
                    await page.keyboard.press("Control+A")
                    await page.wait_for_timeout(300)
                    
                    # Cola o conteúdo (Ctrl+V)
                    print("  - Colando conteúdo...")
                    await page.keyboard.press("Control+V")
                    await page.wait_for_timeout(2000)
                    
                    print("✓ Conteúdo colado com sucesso")
                    
                    # Agora fecha o Word (Alt+F4) e confirma fechar sem salvar
                    print("  - Fechando Word...")
                    subprocess.run([
                        'powershell',
                        '-Command',
                        '''
                        Add-Type -AssemblyName System.Windows.Forms
                        # Fecha o Word
                        [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
                        Start-Sleep -Milliseconds 1000
                        # Se pergunta sobre salvar, pressiona N (não)
                        [System.Windows.Forms.SendKeys]::SendWait("n")
                        Start-Sleep -Milliseconds 500
                        '''
                    ], check=True, capture_output=True)
                    
                    print("✓ Word fechado")
                    await page.wait_for_timeout(1000)
                    
                except Exception as e:
                    print(f"✗ Erro ao colar conteúdo: {e}")
                    await page.screenshot(path="erro_colar_conteudo.png")
                    raise Exception("Não foi possível colar o conteúdo")
                
                # ============================================================
                # ETAPA 7: SALVAR E FECHAR JANELA DE EDIÇÃO
                # ============================================================
                print("\n[ETAPA 7] Salvando e fechando janela de edição...")
                try:
                    print("  - Localizando botão Salvar...")
                    botao_salvar = page.get_by_role("button", name="Salvar")
                    await botao_salvar.wait_for(timeout=5000)
                    print("  - Clicando em Salvar...")
                    await botao_salvar.click()
                    await page.wait_for_timeout(1500)
                    
                    print("✓ Documento salvo")
                    print("  - Fechando janela de edição...")
                    await page.close()
                    
                    # Volta para página original (primeira página do contexto)
                    paginas = page.context.pages
                    if paginas:
                        page = paginas[0]
                        print("✓ Voltado para janela principal")
                    else:
                        raise Exception("Nenhuma página disponível")
                    
                except Exception as e:
                    print(f"✗ Erro ao salvar/fechar: {e}")
                    await page.screenshot(path="erro_salvar_fechar.png")
                    raise Exception("Não foi possível salvar ou fechar a janela")
                
                # ============================================================
                # ETAPA 8: ASSINAR DOCUMENTO
                # ============================================================
                print("\n[ETAPA 8] Assinando documento...")
                try:
                    print("  - Localizando link 'Assinar Documento'...")
                    iframe_conteudo = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    await iframe_conteudo.wait_for(timeout=10000)
                    frame_conteudo = iframe_conteudo.content_frame
                    
                    link_assinar = frame_conteudo.get_by_role("link", name="Assinar Documento")
                    await link_assinar.wait_for(timeout=5000)
                    
                    # Clica em "Assinar Documento"
                    print("  - Clicando em 'Assinar Documento'...")
                    await link_assinar.click()
                    await page.wait_for_timeout(2000)  # Aguarda modal abrir
                    
                    # Copia a senha para o clipboard
                    print("  - Copiando senha para clipboard...")
                    copiar_para_clipboard(self.senha.get())
                    await page.wait_for_timeout(500)
                    
                    # Cola a senha (Ctrl+V)
                    print("  - Colando senha no modal...")
                    await page.keyboard.press("Control+V")
                    await page.wait_for_timeout(1000)
                    
                    # Clica no botão "Assinar" dentro do modal
                    print("  - Clicando em 'Assinar' no modal...")
                    iframe_modal = page.locator('iframe[name="modal-frame"]')
                    await iframe_modal.wait_for(timeout=5000)
                    frame_modal = iframe_modal.content_frame
                    
                    botao_assinar = frame_modal.get_by_role("button", name="Assinar")
                    await botao_assinar.wait_for(timeout=5000)
                    await botao_assinar.click()
                    
                    print("✓ Documento assinado")
                    await page.wait_for_timeout(2000)
                    
                except Exception as e:
                    print(f"✗ Erro ao assinar: {e}")
                    await page.screenshot(path="erro_assinar.png")
                    raise Exception("Não foi possível assinar o documento")
                
                # ============================================================
                # ETAPA 9: INCLUIR NOVO DOCUMENTO
                # ============================================================
                print("\n[ETAPA 9] Incluindo novo documento...")
                try:
                    # Pausa maior antes de clicar
                    await page.wait_for_timeout(5000)
                    
                    iframe_conteudo = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    frame_conteudo = iframe_conteudo.content_frame
                    
                    link_incluir = frame_conteudo.get_by_role("link", name="Incluir Documento")
                    await link_incluir.wait_for(timeout=5000)
                    await link_incluir.click()
                    print("✓ Link 'Incluir Documento' clicado")
                    
                    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                    await page.wait_for_timeout(2000)
                    
                except Exception as e:
                    print(f"✗ Erro ao clicar em 'Incluir Documento': {e}")
                    await page.screenshot(path="erro_incluir_doc.png")
                    raise Exception("Não foi possível clicar em 'Incluir Documento'")
                
                # ============================================================
                # ETAPA 10: PREENCHER FORMULÁRIO DO NOVO DOCUMENTO
                # ============================================================
                print("\n[ETAPA 10] Preenchendo formulário do novo documento...")
                try:
                    iframe_viz1 = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    frame_viz1 = iframe_viz1.content_frame
                    
                    iframe_viz2 = frame_viz1.locator('iframe[name="ifrVisualizacao"]')
                    frame_viz2 = iframe_viz2.content_frame
                    
                    # Clica em "Externo" (como link)
                    print("  - Clicando em 'Externo'...")
                    link_externo = frame_viz2.get_by_role("link", name="Externo")
                    await link_externo.click()
                    await page.wait_for_timeout(1000)
                    
                    # Seleciona tipo de documento (Relatório Técnico = 735)
                    print("  - Selecionando 'Relatório Técnico'...")
                    select_tipo_doc = frame_viz2.get_by_label("Tipo do Documento:")
                    await select_tipo_doc.select_option("735")
                    await page.wait_for_timeout(500)
                    
                    # Clica no botão de data
                    print("  - Clicando no seletor de data...")
                    btn_data = frame_viz2.get_by_role("img", name="Selecionar Data")
                    await btn_data.click()
                    await page.wait_for_timeout(1000)
                    
                    # Clica no dia de hoje
                    dia_hoje = str(datetime.now().day)
                    print(f"  - Selecionando dia {dia_hoje}...")
                    cell_data = frame_viz2.get_by_role("cell", name=dia_hoje, exact=True)
                    await cell_data.click()
                    await page.wait_for_timeout(500)
                    
                    # Nome na Árvore
                    print("  - Inserindo nome na árvore...")
                    campo_nome = frame_viz2.get_by_role("textbox", name="Nome na Árvore:")
                    await campo_nome.click()
                    await page.wait_for_timeout(300)
                    await campo_nome.fill("")  # Limpa
                    await page.wait_for_timeout(200)
                    nome_arvore = numero_relatorio
                    await campo_nome.fill(nome_arvore)
                    await page.wait_for_timeout(500)
                    
                    # Formato - Nato-digital
                    print("  - Selecionando formato 'Nato-digital'...")
                    label_nato = frame_viz2.locator("#divOptNato > .infraRadioDiv > .infraRadioLabel")
                    await label_nato.click()
                    await page.wait_for_timeout(500)
                    
                    # Nível de Acesso - Restrito
                    print("  - Selecionando nível 'Restrito'...")
                    label_restrito = frame_viz2.locator("#divOptRestrito > .infraRadioDiv > .infraRadioLabel")
                    await label_restrito.click()
                    await page.wait_for_timeout(500)
                    
                    # Hipótese Legal (ID: 61)
                    print("  - Selecionando hipótese legal...")
                    select_hipotese = frame_viz2.get_by_label("Hipótese Legal:")
                    await select_hipotese.select_option("61")
                    await page.wait_for_timeout(500)
                    
                    print("✓ Formulário preenchido")
                    
                except Exception as e:
                    print(f"✗ Erro ao preencher formulário: {e}")
                    await page.screenshot(path="erro_preencher_form.png")
                    raise Exception("Erro ao preencher formulário do novo documento")
                
                # ============================================================
                # ETAPA 11: ANEXAR ARQUIVO PDF
                # ============================================================
                print("\n[ETAPA 11] Anexando arquivo PDF...")
                try:
                    iframe_viz1 = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    frame_viz1 = iframe_viz1.content_frame
                    
                    iframe_viz2 = frame_viz1.locator('iframe[name="ifrVisualizacao"]')
                    frame_viz2 = iframe_viz2.content_frame
                    
                    # Clica em "Anexar Arquivo..."
                    print("  - Clicando em 'Anexar Arquivo...'")
                    btn_anexar = frame_viz2.get_by_text("Anexar Arquivo...")
                    await btn_anexar.wait_for(timeout=5000)
                    await btn_anexar.click()
                    await page.wait_for_timeout(1000)
                    
                    # Usa set_input_files para selecionar o arquivo
                    print(f"  - Anexando arquivo: {Path(caminho_pdf).name}")
                    await btn_anexar.set_input_files(caminho_pdf)
                    await page.wait_for_timeout(1000)

                    # Pressiona Esc para fechar qualquer diálogo residual
                    print("  - Pressionando Esc...")
                    await page.keyboard.press("Escape")
                    await page.wait_for_timeout(1000)

                    print("✓ Arquivo PDF anexado")
                    
                except Exception as e:
                    print(f"✗ Erro ao anexar arquivo: {e}")
                    await page.screenshot(path="erro_anexar_pdf.png")
                    raise Exception("Não foi possível anexar o arquivo PDF")
                
                # ============================================================
                # ETAPA 12: SALVAR NOVO DOCUMENTO
                # ============================================================
                print("\n[ETAPA 12] Salvando novo documento...")
                try:
                    iframe_viz1 = page.locator('iframe[name="ifrConteudoVisualizacao"]')
                    frame_viz1 = iframe_viz1.content_frame
                    
                    iframe_viz2 = frame_viz1.locator('iframe[name="ifrVisualizacao"]')
                    frame_viz2 = iframe_viz2.content_frame
                    
                    # Clica no botão Salvar
                    print("  - Clicando em 'Salvar'...")
                    btn_salvar_doc = frame_viz2.locator("#divInfraBarraComandosInferior").get_by_role("button", name="Salvar")
                    await btn_salvar_doc.wait_for(timeout=5000)
                    await btn_salvar_doc.click()
                    
                    print("✓ Novo documento salvo")
                    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                    await page.wait_for_timeout(2000)
                    
                except Exception as e:
                    print(f"✗ Erro ao salvar novo documento: {e}")
                    await page.screenshot(path="erro_salvar_novo_doc.png")
                    raise Exception("Não foi possível salvar o novo documento")
                
                print("\n" + "="*70)
                print("✓✓✓ AUTOMAÇÃO CONCLUÍDA COM SUCESSO! ✓✓✓")
                print("="*70)
                print("\n⚠️  Janela do navegador será mantida aberta para consulta.")
                print("Feche manualmente quando terminar de usar.\n")
                
            except Exception as e:
                print(f"✗ Erro: {str(e)}")
                raise
            finally:
                # Mantém navegador aberto - comentado para não fechar após automação
                # print("Fechando navegador...")
                # await browser.close()
                pass


def main():
    root = tk.Tk()
    app = AutomacaoSEI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
