from modules.relatorio_suportes_tr import formata_suporte
from modules.relatorio_andamento import formata_andamentos
from modules.relatorio_geral import formata_geral

print("---------------------------------------------")
print("Formatador de Relatórios SJ - v2.7 - Escolha:")
print("---------------------------------------------")
print("1 - Andamentos                              *")
print("2 - Suportes TR                             *")
print("3 - Suportes Geral (beta)                   *")
print("*********************************************")

opcao = 0
while opcao != 1 and opcao != 2 and opcao !=3 and opcao !=4:   
    opcao = int(input("Opção: "))
    print("-----------------------------------------------------")
    if opcao != 1 and opcao != 2 and opcao !=3 and opcao !=4:
        print("Opção inválida!")

if opcao == 1:
    print("Relatório de Andamentos gerado na pasta 'excel files'.")
    formata_andamentos()
if opcao == 2:
    print("Relatório de Suportes TR gerado na pasta 'excel files'.")
    formata_suporte()
if opcao == 3:
    print("Relatório de Suportes Geral gerado na pasta 'excel files'.")
    formata_geral()