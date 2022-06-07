from classes import *


def main():
    OutrasFuncoes()
    cliente = Clientes()
    cliente.criar_cliente()
    cliente.listarClientes()
    cliente.selecionar_ano()
    cliente.selecionar_mes()
    cliente.selecionar_aliquota()
    ChromeAuto.baixar_nota()
    navegador = ChromeAuto()
    navegador.abrir_site()
    navegador.gerar_notas(cliente)
    navegador.encerrar()
    

if __name__ == '__main__':
    main()
