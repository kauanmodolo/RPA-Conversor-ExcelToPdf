import fitz
import re
from pathlib import Path
from typing import Dict, Optional

class Pdf_Manager:
    def __init__(self):
        """
        Inicializa o gerenciador de PDF com padrões para:
        - Número do contrato
        - Valor total (formato monetário)
        - Conceito (texto descritivo)
        """
        self.padroes = {
            "contrato": {
                "rotulo": "nº do contrato",
                "regex": re.compile(r'\b\d{8,12}\b')
            },
            "valor_total": {
                "rotulo": "valor total",
                    "regex": re.compile(r'(?i)(?:R\$)?\s*(\d{1,3}(?:\.?\d{3})*,\d{2})')
            },
            "conceito": {
                "rotulo": "conceito",
                "regex": re.compile(r'(?i)conceito\s*\d+\s*(.+)')
            }
        }

    def encontrar_dados(self, caminho_pdf: str) -> Dict:
        """
        Busca em um PDF:
        - Número do contrato
        - Valor total (R$)
        - Conceito (descrição textual)

        Args:
            caminho_pdf (str): Caminho para o arquivo PDF.

        Returns:
            Dict: Dicionário com:
                - 'arquivo': Nome do arquivo
                - 'contrato_principal': Número do contrato
                - 'valor_total': Valor monetário (float)
                - 'conceito': Texto descritivo
                - 'metodo': Método de detecção usado

        Raises:
            FileNotFoundError: Se o arquivo não existir
            ValueError: Se o contrato não for encontrado
        """
        caminho_pdf = Path(caminho_pdf)
        if not caminho_pdf.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho_pdf}")

        resultados = {
            "arquivo": caminho_pdf.name,
            "contrato_principal": None,
            "valor_total": None,
            "conceito": None,
            "metodo": None
        }

        with fitz.open(caminho_pdf) as doc:
            # Busca por rotulos primeiro
            for pagina in doc:
                texto = pagina.get_text("text")
                
                if not resultados["contrato_principal"]:
                    resultados["contrato_principal"] = self._buscar_por_rotulo(texto, "contrato")
                    if resultados["contrato_principal"]:
                        print("🔎 Contrato encontrado pela função: _buscar_por_rotulo")
                
                if not resultados["valor_total"]:
                    resultados["valor_total"] = self._buscar_por_rotulo(texto, "valor_total")
                
                if not resultados["conceito"]:
                    resultados["conceito"] = self._buscar_por_rotulo(texto, "conceito")

            # Fallback para contrato se não encontrado
            if not resultados["contrato_principal"]:
                resultados["contrato_principal"] = self._buscar_generico(doc, "contrato")
                resultados["metodo"] = "genérico"
                if resultados["contrato_principal"]:
                    print("🔎 Contrato encontrado pela função: _buscar_generico")
            else:
                resultados["metodo"] = "rotulo especifico"

        if not resultados["contrato_principal"]:
            raise ValueError(f"Nenhum contrato encontrado em {caminho_pdf.name}")

        return resultados

    def _buscar_por_rotulo(self, texto: str, tipo: str) -> Optional[str]:
        """
        Busca informações baseadas em rótulos específicos.

        Args:
            texto (str): Texto extraído da página do PDF
            tipo (str): Tipo de dado a buscar ('contrato', 'valor_total', 'conceito')

        Returns:
            Optional[str]: Valor encontrado ou None
        """
        config = self.padroes[tipo]
        linhas = texto.split('\n')
        
        for i, linha in enumerate(linhas):
            linha_limpa = linha.strip().lower()
            if config["rotulo"].lower() in linha_limpa:
                print(f"🏷️  Rótulo '{config['rotulo']}' encontrado na linha {i+1}")
                
                # Para conceito, busca só na linha 2 de baixo (NÃO usa regex)
                if tipo == "conceito":
                    if i + 2 < len(linhas):
                        conceito_linha = linhas[i + 2].strip()
                        if conceito_linha:
                            print(f"✅ Valor encontrado na linha {i+2}: {conceito_linha}")
                            return conceito_linha
                    return None

                # Para outros tipos, usa regex na linha do rótulo
                match = config["regex"].search(linha)
                if match:
                    print(f"✅ Valor encontrado na linha {i+1}: {linha.strip()}")
                    return self._processar_valor(tipo, match)
                
                # Se não encontrar, pega a próxima linha não vazia
                for j in range(i+1, len(linhas)):
                    if linhas[j].strip():
                        if tipo == "valor_total":
                            valor_str = linhas[j].strip().replace(".", "").replace(",", ".").replace("R$", "").strip()
                            try:
                                print(f"✅ Valor encontrado na linha {j+1}: {linhas[j].strip()}")
                                return float(valor_str)
                            except Exception:
                                print(f"✅ Valor encontrado na linha {j+1}: {linhas[j].strip()}")
                                return linhas[j].strip()
                        else:
                            print(f"✅ Valor encontrado na linha {j+1}: {linhas[j].strip()}")
                            return linhas[j].strip()
        return None


    
    def _processar_valor(self, tipo: str, match: re.Match) -> Optional[str]:
        """
        Formata valores extraídos conforme o tipo.

        Args:
            tipo (str): Tipo de dado ('contrato', 'valor_total', 'conceito')
            match (re.Match): Objeto Match do regex

        Returns:
            Valor formatado (int, float, str) ou None
        """
        valor = match.group(1) if match.groups() else match.group()
        print(f"🔧 Valor bruto encontrado ({tipo}): {valor}")  # Debug
        
        if tipo == "contrato":
            return int(valor)
        elif tipo == "valor_total":
            return float(valor.replace(".", "").replace(",", "."))
        elif tipo == "conceito":
            return valor.strip()
        return None

    def _buscar_generico(self, doc, tipo: str) -> Optional[str]:
        """
        Busca genérica em todo o documento (fallback).

        Args:
            doc: Documento PDF aberto
            tipo (str): Tipo de dado a buscar

        Returns:
            Optional[str]: Valor encontrado ou None
        """
        config = self.padroes[tipo]
        for pagina in doc:
            texto = pagina.get_text("text")
            match = config["regex"].search(texto)
            if match:
                return self._processar_valor(tipo, match)
        return None