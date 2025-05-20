from dataclasses import dataclass

@dataclass
class NotaFiscalModel:
    CONTRATO: int
    CODIGO_PARCEIRO: int
    TIPO_DE_NEGOCIACAO: int
    TIPO_OPERACAO: int
    CENTRO_RESULTADO: int
    PROJETO: int
    NATUREZA: int
    CIDADE: int
    CIDADE_SERVICO: int
    PRODUTO: int