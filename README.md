# AlterarMultiplosDadosNoBD
Projeto feito em VBA, utilização principal alterar dados em grande escala. 
Basicamente tive a ideia de criar esse projeto quando estava na empresa que eu trabalho e recebi uma demanda onde tinha uma planilha de EXCEL com vários dados (uns 400) e basicamente a única coisa que eu teria que fazer seria alterar os valores de um atributo do banco. Ao invés de sair colando a chave primária de cada elemento a ser alterado resolvi criar essa macro, onde ele pega todos os dados em massa e coloca a estrutura de um bloco de notas, que é criado no Desktop como um arquivo .txt denominado script. Após ser criado basta copiar e colar esse script no banco e pronto, sem dor de cabeça, rápido e eficaz!

O fluxo de funcionamento basicamente é: 

1º Informa a coluna do excel.
2º Informa a linha inicial a ser percorrida. 
3º Tabela do banco onde vai ocorrer as alterações
4º O atributo do banco que vai sofrer alteração
5º O novo valor que esse atributo vai receber
6º E por fim a chave primária da base de dados.



