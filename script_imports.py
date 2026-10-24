import argparse
import ast
import json
import re
from pathlib import Path
from typing import Iterable, List, Set
from openpyxl import Workbook
from openpyxl.styles import Font


# Extensões suportadas por padrão na varredura automática
EXTENSOES_PADRAO = {".py", ".ipynb", ".json", ".txt", ".sql", ".scala", ".r"}

# Diretórios comuns que devem ser ignorados durante a busca
DIRETORIOS_IGNORADOS = {
    ".git",
    "__pycache__",
    ".ipynb_checkpoints",
    "venv",
    ".venv",
    "env",
    "node_modules",
    ".idea",
    ".vscode",
}

# Expressão regular para capturar linhas de importação (usada como fallback ou para arquivos não-Python puros)
PADRAO_IMPORT_REGEX = re.compile(
    r"^\s*(?:from\s+[\w\.]+\s+import\s+[^\n#;]+|import\s+[^\n#;]+)",
    re.MULTILINE,
)


def extrair_imports_via_ast(codigo: str) -> Set[str]:
    """
    Extrai declarações de importação de código Python válido usando a árvore sintática (AST).
    Ignora comentários e strings multiline automaticamente.
    """
    imports_encontrados: Set[str] = set()
    try:
        arvore = ast.parse(codigo)
        for no in ast.walk(arvore):
            if isinstance(no, ast.Import):
                nomes = ", ".join(
                    f"{alias.name} as {alias.asname}" if alias.asname else alias.name
                    for alias in no.names
                )
                imports_encontrados.add(f"import {nomes}")
            elif isinstance(no, ast.ImportFrom):
                modulo = "." * no.level + (no.module or "")
                nomes = ", ".join(
                    f"{alias.name} as {alias.asname}" if alias.asname else alias.name
                    for alias in no.names
                )
                imports_encontrados.add(f"from {modulo} import {nomes}")
    except SyntaxError:
        # Se o bloco contiver comandos mágicos (%pyspark, %sql, !pip) ou sintaxe incompleta,
        # recorre à extração via expressão regular.
        return extrair_imports_via_regex(codigo)

    return imports_encontrados


def extrair_imports_via_regex(texto: str) -> Set[str]:
    """
    Extrai declarações de importação usando Regex.
    Útil para células do Apache Zeppelin (com diretivas %pyspark), trechos parciais ou outros formatos.
    """
    imports_encontrados: Set[str] = set()
    for correspondencia in PADRAO_IMPORT_REGEX.findall(texto):
        linha_limpa = " ".join(correspondencia.strip().split())
        imports_encontrados.add(linha_limpa)
    return imports_encontrados


def extrair_textos_de_json(dados: dict) -> List[str]:
    """
    Extrai blocos de código de estruturas JSON conhecidas:
    - Notebooks Apache Zeppelin ('paragraphs' -> 'text')
    - Jupyter Notebooks ('cells' -> 'source')
    """
    blocos_texto: List[str] = []

    if not isinstance(dados, dict):
        return blocos_texto

    # Formato Apache Zeppelin (ex: note.json)
    if "paragraphs" in dados and isinstance(dados["paragraphs"], list):
        for paragrafo in dados["paragraphs"]:
            if isinstance(paragrafo, dict) and "text" in paragrafo:
                texto = paragrafo["text"]
                if isinstance(texto, str):
                    blocos_texto.append(texto)

    # Formato Jupyter Notebook (.ipynb)
    if "cells" in dados and isinstance(dados["cells"], list):
        for celula in dados["cells"]:
            if isinstance(celula, dict) and celula.get("cell_type") in ("code", None):
                fonte = celula.get("source", [])
                if isinstance(fonte, list):
                    blocos_texto.append("".join(fonte))
                elif isinstance(fonte, str):
                    blocos_texto.append(fonte)

    return blocos_texto


def processar_arquivo(caminho: Path, usar_ast: bool = True) -> Set[str]:
    """
    Lê qualquer arquivo suportado (.py, .ipynb, .json ou texto plano) e retorna os imports encontrados.
    """
    imports_arquivo: Set[str] = set()

    try:
        conteudo_bruto = caminho.read_text(encoding="utf-8-sig", errors="replace")
    except OSError as erro:
        print(f"[AVISO] Não foi possível ler '{caminho}': {erro}")
        return imports_arquivo

    sufixo = caminho.suffix.lower()

    # Tratamento especial para arquivos estruturados em JSON (.json e .ipynb)
    if sufixo in {".json", ".ipynb"}:
        try:
            dados_json = json.loads(conteudo_bruto)
            blocos = extrair_textos_de_json(dados_json)
        except json.JSONDecodeError:
            blocos = [conteudo_bruto]
    else:
        blocos = [conteudo_bruto]

    for bloco in blocos:
        if usar_ast and sufixo in {".py", ".ipynb"}:
            imports_arquivo.update(extrair_imports_via_ast(bloco))
        else:
            imports_arquivo.update(extrair_imports_via_regex(bloco))

    return imports_arquivo


def buscar_arquivos(
    diretorio_raiz: Path,
    padroes_nome: Iterable[str] = ("*",),
    extensoes: Set[str] = EXTENSOES_PADRAO,
) -> Iterable[Path]:
    """
    Percorre recursivamente o diretório raiz filtrando pelos padrões ou extensões desejadas.
    """
    for caminho in diretorio_raiz.rglob("*"):
        if not caminho.is_file():
            continue

        # Ignora pastas de ambiente virtual, cache ou controle de versão
        if any(parte in DIRETORIOS_IGNORADOS for parte in caminho.parts):
            continue

        # Se o usuário especificou padrões exatos (ex: 'note.json', '*.py'), valida o match
        if padroes_nome and padroes_nome != ("*",):
            if any(caminho.match(padrao) for padrao in padroes_nome):
                yield caminho
        elif caminho.suffix.lower() in extensoes:
            yield caminho


def exportar_para_excel(
    imports_unicos: List[str],
    registros_detalhados: List[tuple],
    caminho_saida: Path,
) -> None:
    """
    Salva a lista consolidada de imports únicos e uma aba opcional de mapeamento por arquivo em Excel.
    """
    workbook = Workbook()

    # Aba 1: Lista consolidada sem duplicatas
    aba_resumo = workbook.active
    aba_resumo.title = "Imports Únicos"
    aba_resumo.cell(row=1, column=1, value="Declaração de Import").font = Font(bold=True)

    for indice, item in enumerate(imports_unicos, start=2):
        aba_resumo.cell(row=indice, column=1, value=item)

    aba_resumo.column_dimensions["A"].width = 60

    # Aba 2: Rastreabilidade (qual arquivo contém qual import)
    aba_detalhes = workbook.create_sheet(title="Por Arquivo")
    aba_detalhes.cell(row=1, column=1, value="Arquivo").font = Font(bold=True)
    aba_detalhes.cell(row=1, column=2, value="Declaração de Import").font = Font(bold=True)

    for indice, (arquivo_relativo, declaracao) in enumerate(registros_detalhados, start=2):
        aba_detalhes.cell(row=indice, column=1, value=str(arquivo_relativo))
        aba_detalhes.cell(row=indice, column=2, value=declaracao)

    aba_detalhes.column_dimensions["A"].width = 50
    aba_detalhes.column_dimensions["B"].width = 60

    workbook.save(filename=caminho_saida)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extrai declarações de importação de qualquer arquivo (.py, .ipynb, .json, etc.) e exporta para Excel."
    )
    parser.add_argument(
        "-d",
        "--diretorio",
        type=Path,
        default=Path.cwd(),
        help="Diretório raiz para varredura (padrão: diretório atual).",
    )
    parser.add_argument(
        "-f",
        "--filtro",
        nargs="+",
        default=["*"],
        help="Filtro de nomes ou padrões de arquivo (ex: 'note.json' '*.py' '*.ipynb'). Padrão: todos os suportados.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("lista_imports.xlsx"),
        help="Nome/caminho do arquivo Excel de saída (padrão: lista_imports.xlsx).",
    )
    parser.add_argument(
        "--apenas-regex",
        action="store_true",
        help="Força o uso apenas de Regex em vez de AST para todos os arquivos.",
    )

    args = parser.parse_args()
    diretorio_raiz = args.diretorio.resolve()

    if not diretorio_raiz.exists():
        print(f"[ERRO] O diretório '{diretorio_raiz}' não existe.")
        return

    todos_imports: Set[str] = set()
    registros_detalhados: List[tuple] = []
    total_arquivos = 0
    total_ocorrencias = 0

    for caminho_arquivo in buscar_arquivos(diretorio_raiz, padroes_nome=args.filtro):
        total_arquivos += 1
        imports_no_arquivo = processar_arquivo(
            caminho_arquivo, usar_ast=not args.apenas_regex
        )
        total_ocorrencias += len(imports_no_arquivo)
        todos_imports.update(imports_no_arquivo)

        caminho_relativo = caminho_arquivo.relative_to(diretorio_raiz)
        for imp in sorted(imports_no_arquivo):
            registros_detalhados.append((caminho_relativo, imp))

    imports_ordenados = sorted(todos_imports)

    print(f"Arquivos analisados: {total_arquivos}")
    print(f"Total de imports encontrados (com repetições entre arquivos): {total_ocorrencias}")
    print(f"Total de imports únicos: {len(imports_ordenados)}")

    exportar_para_excel(imports_ordenados, registros_detalhados, args.output)
    print(f"Planilha salva em: {args.output.resolve()}")


if __name__ == "__main__":
    main()
