from modules.relatorio_suportes_tr import formata_suporte
from modules.relatorio_andamentos import formata_andamentos
from modules.relatorio_geral import formata_geral

print("-----------------------------------------------------")
print("Bem vindo ao formatador de relatórios do SJ! Escolha:")
print("-----------------------------------------------------")
print("1 - Formatar um relatório de Andamentos             *")
print("2 - Formatar um relatório de Suportes TR            *")
print("3 - Formatar um relatório Geral (beta)               ")
print("4 - Todos                                           *")
print("*****************************************************")

opcao = 0
while opcao != 1 and opcao != 2 and opcao !=3 and opcao !=4:   
    opcao = int(input("Opção: "))
    print("-----------------------------------------------------")
    if opcao != 1 and opcao != 2 and opcao !=3 and opcao !=4:
        print("Opção inválida!")

if opcao == 1:
    print("Relatório gerado: Andamentos")
    formata_andamentos()
if opcao == 2:
    print("Relatório gerado: Suportes TR")
    formata_suporte()
if opcao == 3:
    print("Relatório gerado: Geral")
    formata_geral()
if opcao == 4:
    print("Relatórios gerados:")
    print()
    print("Andamentos")
    print("Suportes TR")
    formata_andamentos()
    formata_suporte()
