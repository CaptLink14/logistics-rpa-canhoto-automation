import fitz
import time
import shutil
import os
import pandas as pd
import re
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

# ==========================================
# ⚙️ CONFIGURAÇÕES GERAIS
# ==========================================
# Caminho da Planilha
CAMINHO_EXCEL_MESTRE = r"dados/planilha_exemplo.xlsx" #Coloque o caminho do arquivo Excel aqui

# Pasta onde serão criadas as subpastas
PASTA_RAIZ_REDE = r"saida/canhotos" # Coloque o caminho que deseja onde vai ser criada as pastas para cada transportadora

# Nomes das colunas no Excel
COL_NF = "NF" # Nome da coluna que contém o número da nota fiscal (ex: "NF", "Nota Fiscal", etc.)
COL_SS = "SS" # Nome da coluna que contém o número do Pedido (ex: "ID_pedido", "Pedido", etc.)
COL_SIGLA = "Sigla Cliente" # Nome da coluna que contém a identificação do cliente (ex: "Sigla", "Cliente", etc.)
COL_TRANSP = "Transportadora" # Nome da coluna que contém o nome da transportadora (ex: "Transportadora", "Transp", etc.)

# ==============================================================================
# 1. FUNÇÃO: Transportadora 1 (SISTEMA ESL CLOUD)
# ==============================================================================
def processar_transportadora_1(playwright, lista_dados, pasta_destino): #Ajusta o nome da função para indicar a transportadora correta
    print(f"\n>>> INICIANDO TRANSPORTADORA 1 ({len(lista_dados)} notas) <<<")
    
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    
    sucessos = []
    erros = []

    try:
        print("🔑 Acessando Transportadora 1 ...")
        page.goto("https://sitedatransportadora1.eslcloud.com.br")
        if page.get_by_role("textbox", name="E-mail").is_visible():
            page.get_by_role("textbox", name="E-mail").fill("email@empresa.com.br") # Substitua pelo email real
            page.get_by_role("textbox", name="Senha").fill("Senha") # Substitua pela senha real
            page.get_by_role("button", name="Entrar").click()
            page.wait_for_url("**/dashboard", timeout=15000)
    except Exception as e:
        print(f"❌ Erro Login Transportadora 1 t: {e}") #Ajusta a mensagem para indicar a transportadora correta
        return [], [f"Erro Geral Login: {e}"]

    for i, (nota, ss, sigla) in enumerate(lista_dados):
        ss_limpo = ss.replace('/', '-').replace('\\', '-') if ss else "0000" # Substitui barras por hífens para evitar problemas em nomes de arquivos
        sigla_limpa = sigla.replace('/', '-').replace('\\', '-') if sigla else "CLIENTE" 
        print(f"   [Transportadora 1 {i+1}/{len(lista_dados)}] Nota {nota}...", end="")

        try:
            page.goto("https://sitedatransportadora1.eslcloud.com.br") # Substitua pela URL real
            if page.get_by_role("textbox", name="Nota fiscal").is_visible():
                page.get_by_role("textbox", name="Nota fiscal").fill(nota) 
                page.locator("#submit").click()
                time.sleep(1.5)
            else:
                print(" ❌ Busca sumiu.")
                erros.append(f"{nota} - Busca sumiu")
                continue

            # Lógica (Count > 0 + Scroll)
            botoes_info = page.get_by_title("Informações")
            if botoes_info.count() > 0:
                botoes_info.first.click()
                time.sleep(1)
                page.keyboard.press("PageDown"); page.keyboard.press("PageDown"); time.sleep(0.5)

                botoes_imagem = page.get_by_title("Imagem completa")
                if botoes_imagem.count() > 0:
                    try:
                        with page.expect_download(timeout=30000) as download_info:
                            with page.expect_popup() as popup_info:
                                botoes_imagem.first.click()
                            page_popup = popup_info.value
                        
                        download = download_info.value
                        nome_temp = f"temp_transportadora_1_{nota}.pdf" # Nome temporário para o PDF baixado
                        download.save_as(nome_temp)
                        page_popup.close()
                        
                        # Fecha modal
                        try: page.locator("a").filter(has_text="Fechar").nth(2).click(timeout=2000)
                        except: page.keyboard.press("Escape")

                        if os.path.exists(nome_temp) and os.path.getsize(nome_temp) > 0:
                            doc = fitz.open(nome_temp)
                            imagem = doc[0].get_pixmap(dpi=150)
                            nome_final = f"Canhoto_{nota}_{ss_limpo}_{sigla_limpa}.jpg" # Nome final do arquivo de imagem
                            imagem.save(nome_final)
                            doc.close()
                            shutil.move(nome_final, os.path.join(pasta_destino, nome_final))
                            os.remove(nome_temp)
                            sucessos.append(nota)
                            print(" ✅ Sucesso")
                        else:
                            print(" ❌ Vazio")
                            erros.append(f"{nota} - Arquivo vazio")
                    except Exception as e_down:
                        print(f" ❌ Erro Down: {e_down}")
                        erros.append(f"{nota} - Erro Download")
                        page.keyboard.press("Escape")
                else:
                    print(" ⚠️ Sem botão")
                    erros.append(f"{nota} - Sem botão imagem")
                    page.keyboard.press("Escape")
            else:
                print(" ❌ Não achou")
                erros.append(f"{nota} - Não encontrada")
        except Exception as e:
            print(f" ❌ Erro: {e}")
            erros.append(f"{nota} - Erro Genérico")

    browser.close()
    return sucessos, erros

# ==============================================================================
# 2. FUNÇÃO: TRANSPORTADORA 2 / BRUDAM
# ==============================================================================
def processar_transportadora2(playwright, lista_dados, pasta_destino): #Ajusta o nome da função para indicar a transportadora correta
    print(f"\n>>> INICIANDO TRANSPORTADORA 2 ({len(lista_dados)} notas) <<<")
    
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    sucessos = []
    erros = []

    try:
        print("🔑 Acessando Transportadora 2...")
        page.goto("https://transportadora2.brudam.com.br")
        
        def garantir_login():
            if page.get_by_role("textbox", name="Usuário").is_visible():
                page.get_by_role("textbox", name="Usuário").fill("email@empresa.com.br") # Substitua pelo email real
                page.get_by_role("textbox", name="Senha").fill("Senha") # Substitua pela senha real
                page.get_by_role("button", name=" Acessar o Sistema").click()
                page.wait_for_load_state("networkidle")

        garantir_login()
    except Exception as e:
        print(f"❌ Erro Login Transportadora 2: {e}")
        return [], [f"Erro Geral Login: {e}"]

    for i, (nota, ss, sigla) in enumerate(lista_dados):
        ss_limpo = ss.replace('/', '-').replace('\\', '-') if ss else "S_SS"
        sigla_limpa = sigla.replace('/', '-').replace('\\', '-') if sigla else "CLIENTE"
        print(f"   [Transportadora 2 {i+1}/{len(lista_dados)}] Nota {nota}...", end="") #Ajusta a mensagem para indicar a transportadora correta

        try:
            try:
                page.goto("https://transportadora2.brudam.com.br/site/lista_minuta.php") # Substitua pela URL real
            except:
                time.sleep(5); page.goto("https://transportadora2.brudam.com.br") # Tentativa alternativa caso haja redirecionamento ou mudança de URL
            
            garantir_login()
            
            # Lógica de busca e download
            if page.locator("#busca").is_visible():
                page.locator("#busca").fill(nota)
                page.get_by_role("button", name="Pesquisar").click()
                time.sleep(0.5)
            else:
                erros.append(f"{nota} - Busca sumiu"); continue

            localizador_nota = page.locator("#nomeestranho").get_by_role("cell", name=nota, exact=True)
            if localizador_nota.count() > 0:
                localizador_nota.first.click()
                botao_download = page.locator("td:nth-child(5) > a").last
                
                if botao_download.is_visible(timeout=3000):
                    try:
                        with page.expect_download() as download_info:
                            with page.expect_popup() as page1_info:
                                botao_download.click()
                            page1 = page1_info.value
                        
                        download = download_info.value
                        nome_temp = f"temp_mvt_{nota}.pdf"
                        download.save_as(nome_temp)
                        page1.close(); time.sleep(1.5)
                        
                        if os.path.exists(nome_temp) and os.path.getsize(nome_temp) > 0:
                            doc = fitz.open(nome_temp)
                            imagem = doc[0].get_pixmap(dpi=150)
                            nome_final = f"Canhoto_{nota}_{ss_limpo}_{sigla_limpa}.jpg"
                            imagem.save(nome_final)
                            doc.close()
                            shutil.move(nome_final, os.path.join(pasta_destino, nome_final))
                            os.remove(nome_temp)
                            sucessos.append(nota)
                            print(" ✅ Sucesso")
                        else:
                            print(" ❌ Vazio")
                            erros.append(f"{nota} - Arquivo vazio")
                    except Exception as e_down:
                        print(f" ❌ Erro Down: {e_down}")
                        erros.append(f"{nota} - Erro Download")
                else:
                    print(" ⚠️ Sem anexo")
                    erros.append(f"{nota} - Sem anexo")
            else:
                print(" ❌ Não encontrada")
                erros.append(f"{nota} - Não encontrada")
        except Exception as e:
            print(f" ❌ Erro: {e}")
            erros.append(f"{nota} - Erro Genérico")

    browser.close()
    return sucessos, erros

# ==============================================================================
# 🚀 ORQUESTRADOR PRINCIPAL (MAIN)
# ==============================================================================
def main():
    inicio_total = time.time()
    print(">>> Processo iniciado <<<")
    
    # 1. Lê a Planilha Mestra
    try:
        print(f"📂 Lendo planilha mestra: {CAMINHO_EXCEL_MESTRE}")
        df = pd.read_excel(CAMINHO_EXCEL_MESTRE, dtype=str)
        df.columns = df.columns.str.strip()
        
        # Verifica colunas
        cols_obrigatorias = [COL_NF, COL_SS, COL_SIGLA, COL_TRANSP]
        for col in cols_obrigatorias:
            if col not in df.columns:
                print(f"❌ ERRO: Coluna '{col}' não encontrada na planilha!")
                return

        # Limpeza
        for col in cols_obrigatorias:
            df[col] = df[col].astype(str).str.replace('.0', '', regex=False).str.replace('nan', '', case=False).str.strip()
        
        # Remove linhas vazias
        df = df[df[COL_NF] != ""]
        print(f"✅ Total de linhas carregadas: {len(df)}")
        
    except Exception as e:
        print(f"❌ Erro ao abrir Excel: {e}")
        return

    # 2. Cria estrutura de pastas
    if not os.path.exists(PASTA_RAIZ_REDE):
        try: os.makedirs(PASTA_RAIZ_REDE)
        except: pass

    relatorio_geral = []

    with sync_playwright() as p:

        # --- FILTRO E EXECUÇÃO: Transportadora 1 (ESL CLOUDS) ---
        df_transportadora1 = df[df[COL_TRANSP].str.upper().isin(['Transportadora 1', 'Transp 1'])] #Ajusta os nomes igualmente para o filtro da transportadora correta
        if not df_transportadora1.empty:
            dados = list(zip(df_transportadora1[COL_NF], df_transportadora1[COL_SS], df_transportadora1[COL_SIGLA]))
            pasta = os.path.join(PASTA_RAIZ_REDE, "Canhotos_transportadora_1") #Ajusta o nome da pasta para indicar a transportadora correta
            if not os.path.exists(pasta): os.makedirs(pasta)
            
            suc, err = processar_transportadora_1(p, dados, pasta)
            relatorio_geral.append(f"Transportadora 1: {len(suc)} sucessos, {len(err)} erros.") #Ajusta a mensagem para indicar a transportadora correta

        # --- FILTRO E EXECUÇÃO: TRANSPORTADORA 2 (BRUDAM) ---
        df_transportadora2 = df[df[COL_TRANSP].str.upper().isin(['Transportadora 2', 'Transp 2'])] #Ajusta os nomes igualmente para o filtro da transportadora correta
        if not df_transportadora2.empty: 
            dados = list(zip(df_transportadora2[COL_NF], df_transportadora2[COL_SS], df_transportadora2[COL_SIGLA]))
            pasta = os.path.join(PASTA_RAIZ_REDE, "Canhotos_transportadora_2") #Ajusta o nome da pasta para indicar a transportadora correta
            if not os.path.exists(pasta): os.makedirs(pasta)
            
            suc, err = processar_transportadora2(p, dados, pasta) #Ajusta o nome da função para indicar a transportadora correta
            relatorio_geral.append(f"Transportadora 2: {len(suc)} sucessos, {len(err)} erros.") #Ajusta a mensagem para indicar a transportadora correta

    # 3. Resumo Final
    tempo_total = time.time() - inicio_total
    mins, segs = divmod(tempo_total, 60)
    
    print("\n" + "="*50)
    print(f"🏁 PROCESSO FINALIZADO EM {int(mins)}m {int(segs)}s")
    print("RESUMO:")
    for linha in relatorio_geral:
        print(f" > {linha}")
    print("="*50)

if __name__ == "__main__":
    main()