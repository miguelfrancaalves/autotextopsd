import win32com.client
import pandas as pd
import os
from pathlib import Path
import re
import argparse
import sys
import time
import numpy as np

def limpar_nome(nome):
    caracteres_especiais = r'[<>:"/\|?*]'
    nome_limpo = re.sub(caracteres_especiais, '', str(nome))
    nome_limpo = nome_limpo.strip()
    return nome_limpo

def editar_e_exportar(arquivo_excel=None, nome_camada=None, pasta_saida=None, qualidade=100, formato="PNG", aguardar_enter=True):
    try:
        print("\n=== Iniciando processamento ===")
        
        try:
            ps = win32com.client.Dispatch("Photoshop.Application")
        except Exception:
            print("Erro: Não foi possível conectar ao Photoshop. Verifique se ele está aberto.")
            return False
            
        try:
            doc = ps.ActiveDocument
        except Exception:
            print("Erro: Nenhum documento aberto no Photoshop. Abra um arquivo PSD primeiro.")
            return False
            
        if not arquivo_excel:
            arquivo_excel = 'lista_nomes.xlsx'
        
        if not os.path.exists(arquivo_excel):
            print(f"Erro: Arquivo Excel '{arquivo_excel}' não encontrado.")
            return False
            
        try:
            planilha = pd.read_excel(arquivo_excel)
            
            if 'nome' not in planilha.columns:
                print(f"Erro: O arquivo Excel não contém uma coluna chamada 'nome'.")
                print(f"Colunas encontradas: {', '.join(planilha.columns)}")
                return False
                
            planilha = planilha.dropna(subset=['nome'])
            planilha = planilha[planilha['nome'].astype(str).str.strip() != '']
            planilha = planilha[planilha['nome'].astype(str).str.lower() != 'nan']
            
            if len(planilha) == 0:
                print("Erro: Não foram encontrados nomes válidos no arquivo Excel.")
                return False
                
            print(f"Arquivo Excel carregado: {arquivo_excel} ({len(planilha)} nomes válidos)")
        except Exception as e:
            print(f"Erro ao abrir o arquivo Excel: {str(e)}")
            return False
            
        if not pasta_saida:
            pasta_saida = 'PNG_Exportados'
            
        pasta_export = Path.cwd() / pasta_saida
        try:
            pasta_export.mkdir(exist_ok=True)
            print(f"Pasta principal criada/verificada: {pasta_export}")
        except Exception as e:
            print(f"Erro ao criar pasta principal: {str(e)}")
            return False
            
        if not nome_camada:
            nome_camada = "Alterar Nome"  
        
        camada_encontrada = False
        for layer in doc.ArtLayers:
            if layer.Name == nome_camada:
                camada_encontrada = True
                if layer.Kind != 2:  
                    print(f"Erro: A camada '{nome_camada}' não é uma camada de texto.")
                    return False
                break
                
        if not camada_encontrada:
            print(f"Erro: Camada '{nome_camada}' não encontrada no documento. Camadas disponíveis:")
            for layer in doc.ArtLayers:
                print(f"  - {layer.Name}")
            return False

        total = len(planilha)
        processados = 0
        falhas = 0
        
        print("\nIniciando exportação de arquivos...")
        tempo_inicio = time.time()

        for index, row in planilha.iterrows():
            try:
                nome = str(row['nome']).strip()
                
                if not nome or nome.lower() == 'nan':
                    continue
                    
                nome_limpo = limpar_nome(nome)
                if not nome_limpo:
                    print(f"Aviso: Nome inválido ignorado na linha {index+2}")
                    continue
                
                inicial = nome_limpo[0].upper()
                pasta_letra = pasta_export / inicial
                try:
                    pasta_letra.mkdir(exist_ok=True)
                except Exception as e:
                    print(f"Erro ao criar subpasta {inicial}: {str(e)}")
                    falhas += 1
                    continue

                for layer in doc.ArtLayers:
                    if layer.Name == nome_camada:
                        layer.TextItem.Contents = nome
                        break
                
                options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
                options.Format = 13 
                options.PNG8 = False 
                options.Transparency = True
                options.Quality = qualidade
                
                novo_nome = f"{nome_limpo}.png"
                caminho_png = pasta_letra / novo_nome
                doc.Export(ExportIn=str(caminho_png), ExportAs=2, Options=options)
                
                processados += 1
                progresso = int((processados / total) * 100)
                print(f"[{progresso}%] Exportado: {inicial}/{novo_nome}")
                
            except Exception as e:
                print(f"Erro ao processar nome '{nome}': {str(e)}")
                falhas += 1
                
        tempo_total = time.time() - tempo_inicio
        print(f"\nProcessamento concluído em {tempo_total:.1f} segundos")
        print(f"Total de nomes: {total}")
        print(f"Processados com sucesso: {processados}")
        if falhas > 0:
            print(f"Falhas: {falhas}")
            
        return True
                
    except Exception as e:
        print(f"Erro geral: {str(e)}")
        return False
    finally:
        if aguardar_enter:
            input("\nPressione Enter para sair...")

def mostrar_menu():
    print("\n" + "="*50)
    print("               SKY LABS PHOTOSHOP")
    print("="*50)
    print("Automação de edição de camada de texto e exportação PNG")
    print("="*50)
    print("\nOPÇÕES:")
    print("1. Iniciar processamento com configurações padrão")
    print("2. Configurar o processamento")
    print("3. Verificar configurações atuais")
    print("4. Ajuda")
    print("5. Verificar arquivo Excel")
    print("0. Sair")
    
    opcao = input("\nDigite a opção desejada: ")
    return opcao

def configurar_processamento():
    config = {}
    
    print("\n=== CONFIGURAÇÕES DE PROCESSAMENTO ===")
    
    arquivo_padrao = 'lista_nomes.xlsx'
    resposta = input(f"Nome do arquivo Excel (padrão: {arquivo_padrao}): ")
    config['arquivo_excel'] = resposta if resposta else arquivo_padrao
    
    camada_padrao = "Alterar Nome"
    resposta = input(f"Nome da camada de texto (padrão: {camada_padrao}): ")
    config['nome_camada'] = resposta if resposta else camada_padrao
    
    pasta_padrao = "PNG_Exportados"
    resposta = input(f"Pasta para salvar as imagens (padrão: {pasta_padrao}): ")
    config['pasta_saida'] = resposta if resposta else pasta_padrao
    
    qualidade_padrao = 100
    resposta = input(f"Qualidade da exportação (1-100, padrão: {qualidade_padrao}): ")
    try:
        config['qualidade'] = int(resposta) if resposta else qualidade_padrao
    except ValueError:
        print("Valor inválido. Usando qualidade padrão.")
        config['qualidade'] = qualidade_padrao
    
    print("\nConfiguração concluída!")
    return config

def verificar_excel(arquivo_excel=None):
    print("\n=== VERIFICAÇÃO DO ARQUIVO EXCEL ===")
    
    if not arquivo_excel:
        arquivo_excel = 'lista_nomes.xlsx'
    
    if not os.path.exists(arquivo_excel):
        print(f"Erro: Arquivo Excel '{arquivo_excel}' não encontrado.")
        return
    
    try:
        planilha = pd.read_excel(arquivo_excel)
        
        print(f"Arquivo: {arquivo_excel}")
        print(f"Total de linhas: {len(planilha)}")
        
        if 'nome' not in planilha.columns:
            print("ERRO: O arquivo não contém uma coluna chamada 'nome'")
            print(f"Colunas encontradas: {', '.join(planilha.columns)}")
            return
        
        nulos = planilha['nome'].isna().sum()
        vazios = (planilha['nome'].astype(str).str.strip() == '').sum()
        nan_string = (planilha['nome'].astype(str).str.lower() == 'nan').sum()
        
        planilha_limpa = planilha.dropna(subset=['nome'])
        planilha_limpa = planilha_limpa[planilha_limpa['nome'].astype(str).str.strip() != '']
        planilha_limpa = planilha_limpa[planilha_limpa['nome'].astype(str).str.lower() != 'nan']
        
        validos = len(planilha_limpa)
        
        print(f"\nAnálise da coluna 'nome':")
        print(f"- Valores válidos: {validos}")
        if nulos > 0:
            print(f"- Valores nulos (NaN): {nulos}")
        if vazios > 0:
            print(f"- Strings vazias: {vazios}")
        if nan_string > 0:
            print(f"- Strings 'nan': {nan_string}")
        
        if validos == 0:
            print("\nAVISO: Não há nenhum nome válido para processamento!")
        elif validos < len(planilha):
            print("\nAVISO: Existem valores inválidos que serão ignorados durante o processamento.")
        
        if validos > 0:
            print("\nAmostra dos primeiros 5 nomes válidos:")
            for i, nome in enumerate(planilha_limpa['nome'].head(5)):
                print(f"  {i+1}. {nome}")
                
    except Exception as e:
        print(f"Erro ao analisar o arquivo Excel: {str(e)}")

def mostrar_configuracoes(config):
    print("\n=== CONFIGURAÇÕES ATUAIS ===")
    print(f"Arquivo Excel: {config.get('arquivo_excel', 'lista_nomes.xlsx')}")
    print(f"Nome da camada: {config.get('nome_camada', 'Alterar Nome')}")
    print(f"Pasta de saída: {config.get('pasta_saida', 'PNG_Exportados')}")
    print(f"Qualidade: {config.get('qualidade', 100)}")

def mostrar_ajuda():
    print("\n=== AJUDA ===")
    print("Este programa automatiza a edição de uma camada de texto no Photoshop")
    print("e exporta o resultado como arquivos PNG.")
    print("\nPré-requisitos:")
    print("1. Tenha o Photoshop aberto com um arquivo PSD")
    print("2. Certifique-se que existe uma camada de texto com o nome configurado")
    print("3. Prepare um arquivo Excel com uma coluna 'nome' contendo os textos a inserir")
    print("\nFuncionamento:")
    print("- O programa lerá cada nome do Excel e o inserirá na camada de texto")
    print("- Cada arquivo será exportado como PNG para a pasta configurada")
    print("- Os arquivos serão organizados em subpastas por letra inicial")
    
def modo_linha_comando():
    parser = argparse.ArgumentParser(description='Automação de edição de camada de texto e exportação PNG no Photoshop')
    parser.add_argument('-e', '--excel', help='Caminho para o arquivo Excel com a lista de nomes')
    parser.add_argument('-c', '--camada', help='Nome da camada de texto a ser alterada')
    parser.add_argument('-p', '--pasta', help='Pasta para salvar os arquivos exportados')
    parser.add_argument('-q', '--qualidade', type=int, help='Qualidade da exportação (1-100)')
    parser.add_argument('-s', '--silencioso', action='store_true', help='Modo silencioso (não aguarda Enter no final)')
    
    args = parser.parse_args()
    
    if len(sys.argv) > 1:  
        return editar_e_exportar(
            arquivo_excel=args.excel,
            nome_camada=args.camada,
            pasta_saida=args.pasta,
            qualidade=args.qualidade if args.qualidade else 100,
            aguardar_enter=not args.silencioso
        )
    else:
        return None  

if __name__ == "__main__":
    resultado_cmd = modo_linha_comando()
    
    if resultado_cmd is not None:
        sys.exit(0 if resultado_cmd else 1)
    
    config = {
        'arquivo_excel': 'lista_nomes.xlsx',
        'nome_camada': 'Alterar Nome',
        'pasta_saida': 'PNG_Exportados',
        'qualidade': 100
    }
    
    while True:
        opcao = mostrar_menu()
        
        if opcao == '1':
            editar_e_exportar(**config)
        elif opcao == '2':
            config = configurar_processamento()
        elif opcao == '3':
            mostrar_configuracoes(config)
        elif opcao == '4':
            mostrar_ajuda()
        elif opcao == '5':
            verificar_excel(config.get('arquivo_excel'))
        elif opcao == '0':
            print("\nPrograma finalizado. Obrigado por usar!")
            break
        else:
            print("Opção inválida. Tente novamente.") 