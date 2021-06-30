# Automatização de relatórios diários

## **1. Objetivos**

> Automatizar o processo de filtragem, tratamento e formatação de dados brutos em planilha de relatório a ser entregue diariamente;

> Desenvolver conhecimento e base para futuros projetos de integração entre Python e Excel.

---------------
## **2. Descrição**

O algoritmo foi desenvolvido em sistema Linux, Ubuntu 20.04 versão LTS.

### **2.1 Preparando o ambiente**

- Instalação do Python 3

```
sudo apt install python3
```

OBS: Nem sempre esse comando é necessário, pois na maioria das distribuições linux, Python já vem instalado.

- Instalação do venv (ambiente de desenvolvimento)

Linux:
```
sudo apt update
sudo apt install python3 python3-dev python3-venv
```

- Instalação do pip (para instalar e gerenciar pacotes de bibliotecas python)

```
sudo apt update
sudo apt install pip3
```
- Editor de texto VS Code

   * Download do pacote .deb feito a partir do site oficial: https://code.visualstudio.com/
   * As extensões a serem usadas são referentes aos pacotes instalados, como o Openpyxl.

- Configurando o venv para isolar dependências

Linux:
```
mkdir relatorios_py-xl
cd relatorios_py-xl
python3 -m venv env
```

Windows:
```
mkdir relatorios_py-xl
cd relatorios_py-xl
py -m venv env
```

OBS: a utilização de um ambiente de desenvolvimento, além de organizar melhor o projeto, isola as dependências e bibliotecas para que não haja interferência e problemas com outros pacotes do sistema.

- Ativando o ambiente

Dentro da pasta do projeto:

Linux:
```
source env/bin/activate
```

Windows:
```
.\env\Scripts\activate
```

### **2.2 Bibliotecas**

- Instalação da biblioteca de manipulação de arquivos Excel:

```
pip3 install openpyxl
pip3 install xls2xlsx
```

Para congelar o ambiente com todas as bibliotecas necessárias:

```
pip freeze > requirements.txt
```

### **2.3 Desenvolvendo**

- Exportação de informações de ações abertas no CRM do Shopping Jurídico em arquivo no formato .xls, com filtragem específica:
    + Abertas / Suporte Técnico / 05.05 até dia atual >>> Mostrar 50 itens por página;
- Conversão do arquivo para .xlsx (por enquanto, pelo próprio Excel);
- Copiar arquivo .xlsx para a pasta do ambiente de desenvolvimento;
- Desenvolver algoritmo que:
    + Leia um arquivo .xlsx;
    + Exclua colunas pré-definidas;
    + Exclua linhas que não contém a informação buscada;
    + Formate os dados como tabela (incluindo cores, fontes e filtros);
    + Ajuste a disposição das células (tamanho, largura e altura);
    + Salve um novo arquivo referente à data atual.

----------------
## **3. Utilização**

### **3.1 Método Manual**

> Fazer o download do arquivo .zip disponível no link: https://github.com/shopping-juridico/relatorios_py-xl. Ou executar o comando abaixo:

```
git clone https://github.com/shopping-juridico/relatorios_py-xl.git
```

Desde que o usuário tenha Python instalado no sistema:

1. Descompactar a pasta;
2. Se necessário, colar dentro de "excel files" os arquivos .xlsx com dados brutos;
3. Abrir o terminal dentro do diretório "relatorios_py-xl" e executar os comandos:

Se necessário (em caso de portabilidade entre SO), criar novamente o ambiente virtual com o mesmo nome da pasta:

Linux:
```
python3 -m venv env
```

Windows:
```
py -m venv env
```

Ativar o ambiente:

Linux:
```
source env/bin/activate
```

Windows:
```
.\env\Scripts\activate
```

Instalar as bilbiotecas necessárias:

```
pip3 install openpyxl
pip3 install xls2xlsx
```

Executar o programa:

```
python3 main.py
```

OBS: Se necessário, instalar as bibliotecas pelo comando:

```
pip install -r requirements.txt
```

Após a instalação, o programa pode ser executado sempre a partir do script.

### **3.2 Método arquivo executável**

#### 3.2.1 Sistemas Linux:

- Instalação: Dentro do diretório do projeto, executar o script 'instalação-linux.sh';
- Utilização: Para abrir o software, executar o script 'executar-linux.sh'.

#### 3.2.2 Windows:

- Instalação: Dentro do diretório do projeto, executar o script 'instalação-windows.bat';
- Utilização: Para abrir o software, executar o script 'executar-windows.bat'.

### **3.3 Método interface gráfica**

----------------
## **4. Resultados**

> Criação de um modelo de script de referência ao produzir relatórios automatizados a partir de dados brutos;

> Relatório diário automaticamente formatado para visualização de ações abertas de suportes encaminhados para a Thomson Reuters;

> Relatório diário automaticamente formatado para visualização de horas trabalhadas.
