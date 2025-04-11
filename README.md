# Photoshop - Automação de Texto e Exportação

Ferramenta para automatização no Photoshop que permite editar camadas de texto com base em uma lista de nomes e exportar os resultados como imagens PNG.

## Recursos

- Edição automática de camada de texto no Photoshop
- Leitura de nomes a partir de uma planilha Excel
- Exportação em alta qualidade como arquivos PNG
- Organização de arquivos em pastas por inicial
- Interface interativa com menu de configuração
- Suporte para argumentos via linha de comando
- Instalador automático de dependências
- Verificação e filtragem inteligente de dados Excel

## Pré-requisitos

- Windows com Photoshop instalado
- Python 3.6 ou superior
- Arquivo PSD com camada de texto preparada
- Planilha Excel com coluna 'nome'

## Instalação e Execução

### Método Simples (Recomendado)

Basta executar o arquivo `instalar_e_executar.bat` com duplo clique. Este arquivo:

1. Verifica se o Python está instalado
2. Instala automaticamente todas as dependências necessárias
3. Verifica se o Photoshop está aberto
4. Inicia o programa

### Instalação Manual

Se preferir instalar manualmente:

1. Clone ou baixe este repositório
2. Instale as dependências necessárias executando no prompt de comando:

```
pip install pandas pywin32 openpyxl
```

**Importante**: Se você estiver usando Python instalado pela Microsoft Store, pode ser necessário usar:

```
python -m pip install pandas pywin32 openpyxl
```

## Como usar

### Modo Interface Interativa

1. Execute o arquivo `editar_e_exportar.py`:
   ```
   python editar_e_exportar.py
   ```
2. Abra seu arquivo PSD no Photoshop antes de iniciar
3. No menu, escolha a opção para iniciar o processamento ou configurar

### Modo Linha de Comando

Execute o script com argumentos para configuração:

```
python editar_e_exportar.py --excel caminho/planilha.xlsx --camada "Nome da Camada" --pasta "Pasta_Saida" --qualidade 100
```

#### Argumentos disponíveis:

- `-e, --excel`: Caminho para o arquivo Excel (padrão: lista_nomes.xlsx)
- `-c, --camada`: Nome da camada de texto no Photoshop (padrão: Alterar Nome)
- `-p, --pasta`: Pasta onde serão salvos os arquivos PNG (padrão: PNG_Exportados)
- `-q, --qualidade`: Qualidade da exportação PNG de 1-100 (padrão: 100)
- `-s, --silencioso`: Modo silencioso (não aguarda Enter no final)

## Preparação do arquivo PSD

1. Crie ou abra um arquivo PSD no Photoshop
2. Crie uma camada de texto com o nome "Alterar Nome" (ou o nome que você definir)
3. Posicione o texto conforme desejado

## Preparação do Excel

Crie um arquivo Excel com uma coluna chamada 'nome' contendo os textos que deseja inserir na camada. O arquivo deve estar na mesma pasta do script ou especificado pelo caminho completo.

Exemplo:
| nome         |
|--------------|
| João Silva   |
| Maria Santos |
| Pedro Alves  |

**Importante sobre o Excel:**
- A coluna deve ser nomeada exatamente como 'nome' (minúsculas)
- O programa ignora automaticamente valores vazios ou inválidos (NaN)
- Você pode verificar seu Excel usando a opção 5 no menu principal

## Recursos Avançados

### Verificação do Arquivo Excel

O programa possui uma função específica para verificar seu arquivo Excel antes do processamento:

1. No menu principal, selecione a opção 5 "Verificar arquivo Excel"
2. O programa analisará o arquivo e mostrará:
   - Total de linhas
   - Número de valores válidos
   - Número de valores nulos (NaN)
   - Amostra dos primeiros nomes válidos
   - Avisos sobre problemas potenciais

Esta verificação ajuda a identificar problemas antes de iniciar o processamento, evitando exportações desnecessárias de valores inválidos.

## Estrutura de pastas gerada

```
📁 Pasta do Script
 ┣ 📄 editar_e_exportar.py
 ┣ 📄 instalar_e_executar.bat
 ┣ 📄 lista_nomes.xlsx
 ┗ 📁 PNG_Exportados
    ┣ 📁 J
    │  ┗ 📄 João Silva.png
    ┣ 📁 M
    │  ┗ 📄 Maria Santos.png
    ┗ 📁 P
       ┗ 📄 Pedro Alves.png
```

## Solução de Problemas

### Erro: Exportação de "nan" ou valores vazios

Se o programa estiver exportando arquivos com nomes "nan" (Not a Number) ou outros valores inválidos:

1. Verifique seu arquivo Excel usando a opção 5 do menu
2. Certifique-se de que sua planilha não contém muitas células vazias ou valores NaN
3. Se necessário, limpe sua planilha removendo linhas vazias ou corrija os valores

A versão atual do programa já filtra valores nulos automaticamente, mas é sempre bom manter sua planilha limpa.

### Instalação Automática

Se você encontrar problemas, a maneira mais simples de resolver é usar o instalador automático:

1. Execute `instalar_e_executar.bat` com direitos de administrador
2. O instalador verificará e instalará todas as dependências necessárias

### Erro: No module named 'win32com'

Se você encontrar o erro:

```
ModuleNotFoundError: No module named 'win32com'
```

Siga estas etapas para resolver:

1. Feche todos os prompts de comando e terminais abertos
2. Abra um novo prompt de comando como administrador (clique com o botão direito → Executar como administrador)
3. Execute o comando de instalação específico para o pywin32:

   ```
   pip install pywin32
   ```

4. Se o erro persistir, tente:

   ```
   python -m pip install pywin32
   ```

5. Em alguns casos, pode ser necessário reinstalar com opções específicas:

   ```
   pip uninstall pywin32
   pip install pywin32==305
   ```

6. Após a instalação, execute este comando para garantir que os arquivos estejam registrados corretamente:

   ```
   python -m pywin32_postinstall -install
   ```

### Outros erros comuns

- **Erro de acesso ao Photoshop**: Certifique-se de que o Photoshop está aberto antes de executar o script.
- **Erro ao abrir o Excel**: Verifique se o arquivo Excel está na mesma pasta do script ou especifique o caminho completo.
- **Camada não encontrada**: Confira se o nome da camada no script corresponde ao nome da camada no seu arquivo PSD.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests com melhorias.

## Licença

Este projeto é distribuído como código aberto.

## Autor

Miguel França Alves.
