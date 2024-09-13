from funcoes import create_excel

  romaneios_path = "romaneios.pdf" # pdf com os romaneios de entrada
  gabarito_excel_path = "gabarito.xlsx" # tabela com a primeira coluna de municipios e a segunda coluna com o nome dos estabelecimentos 
  save_path = "resultados_atualizados.xlsx" # saida atualizada com os dados já extraidos
  gerenciais_path = "ConsultasGerenciais.pdf" # consultas do sistema web 
  create_excel(romaneios_path, gerenciais_path, gabarito_excel_path, save_path)
