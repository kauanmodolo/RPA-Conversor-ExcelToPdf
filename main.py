from workers.excel_manager import Excel_Manager
from workers.conversor_pdf_excel import Conversor_Pdf_Excel

if __name__ == "__main__":
    try:
        print("🚀 Iniciando processo de conversão PDF para Excel + Busca")
        
        # 1. Converter PDF para Excel
        pdf_path = "assets/contrato.pdf"
        excel_gerado = "temp_contrato.xlsx"
        
        conversor = Conversor_Pdf_Excel(pdf_path, excel_gerado)
        conversor.executar()
        
        # 2. Extrair dados do Excel gerado
        dados_contrato = conversor.encontrar_dados_no_excel()
        
        print("\n📄 Resultados do Excel gerado:")
        print(f"Contrato: {dados_contrato['contrato_principal']}")
        if dados_contrato['valor_total']:
            print(f"Valor Total: R$ {dados_contrato['valor_total']:.2f}")
        else:
            print("Valor Total: Não encontrado")
        print(f"Conceito: {dados_contrato['conceito']}")

        # 3. Buscar no Excel principal
        excel_manager = Excel_Manager(r"C:\Users\kauan.carrico\OneDrive - Igneo\Área de Trabalho\NOTAS FISCAIS - NATURGY.xlsx")
        excel_manager.carregar_excel()
        
        resultado_excel = excel_manager.buscar_contrato(dados_contrato['contrato_principal'])
        
        print("\n📊 Resultados do Excel principal:")
        for registro in resultado_excel:
            print(registro)

    except Exception as e:
        print(f"\n❌ Erro no processo: {str(e)}")
    finally:
        print("\nProcesso concluído")