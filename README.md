
# Automatizando criação de arquivos através de dados excel .XSLX para arquivos .DOCX com Python

Este projeto foi desenvolvido caso você venha a precisa criar múltiplos arquivos em um determinado formato, especificamente nesse projeto, foi utilizada uma planilha do **excel (.xslx) para o formato (.docx)** ou seja para o famoso e popularmente chamado de "word".

Para ser mais específico, nesse projeto foi utilizado uma planilha excel com **dados fakes** apenas para **fins de teste e aprendizado** e ele automaticamente criou os arquivos em formato **.docx** em um contexto de geração de certificados.

## Autor

- [Francisco Gomes](https://www.linkedin.com/in/fgsj-developer/)


## Aprendizados

Com este projeto foi possível ter uma melhor compreensão do **uso e manipulação** de **arquivos** utilizando algumas bibliotecas. Atrelando também ao entendimento do uso de comando **PIP** para o gerenciamento de dependências para o projeto.

## Código fonte do projeto
![image](https://github.com/user-attachments/assets/dcfac0b9-0616-4e4b-a082-1ecbdbde4590)

- Na raiz do projeto, abra o terminal e instale **pip install openpyxl** e **pip install docxtpl**

> [!NOTE]
> **INFORMAÇÕES IMPORTANTE PARA O SUCESSO DO PROJETO**

### Repare na imagem abaixo, Aqui sobre o arquivo você irá clicar com o botão direito do mouse e selecionar a opção marcada 
![image](https://github.com/user-attachments/assets/01f0bd44-fd20-4db6-b6de-23a2ae588f4d)

![image](https://github.com/user-attachments/assets/aebdeb17-a221-43d0-91a7-f07e8a75f829)

### Ao copiar o caminho do arquivo, você irá colocar nesse trecho aqui, lembrar que remova os colchetes e cole o arquivo.
![image](https://github.com/user-attachments/assets/1ac0d1a3-da21-49cc-a0fc-a23b59545026)
> [!WARNING]
> **ATENÇÃO**

O caminho do path precista está com duas barras invertidas, caso venha uma muito provavelmente irá ocorrer um erro;

### Uso da iteração for em uma tupla
![image](https://github.com/user-attachments/assets/a5df873d-7451-44e5-89be-f5b55434066e)

### Eu comecei a "varredura" a partir do segundo índice que é o [1] e não do primeiro[0], lembrando que em Python, o índice começa em zero. Porquê ao executar o arquivo excel, ele irá gerar uma tupla o qual o índice 0 são os cabeçalhos/colunas o que não tem nenhum dado específico lá.
