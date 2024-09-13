from funcoes import create_excel

romaneios_path = "romaneios.pdf" 
# romaneios_path = "teste.pdf" 
gerenciais_path = "ConsultasGerenciais.pdf"  
gabarito_path = "gabarito.xlsx"
resultados_path = "resultados_atualizados.xlsx"

create_excel(romaneios_path, gerenciais_path, gabarito_path, resultados_path)