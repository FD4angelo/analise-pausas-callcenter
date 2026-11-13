"""
Análise de pausas e tempo logado de operadores de call center.

Cruza duas exportações do discador (eventos de pausa e sessões de login)
e gera um Excel com o resumo por ramal e um painel de quem estourou
o tempo de almoço ou de café.

Uso:
    python analise_pausas.py --pausas dados_exemplo/tempo_pausa.csv \
                             --logins dados_exemplo/tempo_logado.csv \
                             --output relatorio.xlsx
"""

import argparse
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# Limites de pausa em segundos (1 min de tolerância)
LIMITE_ALMOCO = 61 * 60
LIMITE_CAFE = 21 * 60

PAUSAS_PRODUTIVAS = [
    'cliente presencial',
    'fechamento',
    'pendencias',
    'reuniao',
    'feedback',
    'retorno cliente',
    'ligacao manual',
    'atendimento whatsapp',
]
PAUSAS_IMPRODUTIVAS = ['nao disponivel', 'cafe', 'banheiro', 'almoco']


def ler_tabela(caminho: Path) -> pd.DataFrame:
    if caminho.suffix.lower() in {'.xls', '.xlsx'}:
        return pd.read_excel(caminho)
    return pd.read_csv(caminho, sep=';', encoding='utf-8', dtype=str)


def para_segundos(valor) -> int:
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return 0


def hms(segundos) -> str:
    segundos = int(segundos or 0)
    return f'{segundos // 3600:02}:{segundos % 3600 // 60:02}:{segundos % 60:02}'


def normalizar(texto) -> str:
    """Minúsculas e sem acento: 'Almoço' -> 'almoco'."""
    if pd.isna(texto):
        return ''
    texto = unicodedata.normalize('NFKD', str(texto).strip().lower())
    return ''.join(c for c in texto if not unicodedata.combining(c))


def resumir_pausas(df: pd.DataFrame) -> pd.DataFrame:
    for coluna in ('RamalNumero', 'DuracaoEvento', 'RamalStatusDetalhe'):
        if coluna not in df.columns:
            raise ValueError(f'Arquivo de pausas sem a coluna {coluna}.')

    df = df.copy()
    df['RamalNumero'] = df['RamalNumero'].astype(str).str.strip()
    df['duracao'] = df['DuracaoEvento'].apply(para_segundos)
    df['status'] = df['RamalStatusDetalhe'].apply(normalizar)
    df['produtiva'] = df['status'].isin(PAUSAS_PRODUTIVAS)
    df['improdutiva'] = df['status'].isin(PAUSAS_IMPRODUTIVAS)

    def soma_se(mascara):
        return df['duracao'].where(mascara, 0)

    df['seg_produtiva'] = soma_se(df['produtiva'])
    df['seg_improdutiva'] = soma_se(df['improdutiva'])
    df['seg_almoco'] = soma_se(df['status'] == 'almoco')
    df['seg_cafe'] = soma_se(df['status'] == 'cafe')
    df['seg_nao_disponivel'] = soma_se(df['status'] == 'nao disponivel')

    resumo = df.groupby('RamalNumero').agg(
        pausa_produtiva=('seg_produtiva', 'sum'),
        pausa_improdutiva=('seg_improdutiva', 'sum'),
        qtd_produtivas=('produtiva', 'sum'),
        qtd_improdutivas=('improdutiva', 'sum'),
        almoco=('seg_almoco', 'sum'),
        cafe=('seg_cafe', 'sum'),
        nao_disponivel=('seg_nao_disponivel', 'sum'),
        nome_pausa=('Usuario', 'first'),
    )
    return resumo.reset_index()


def resumir_logins(df: pd.DataFrame) -> pd.DataFrame:
    for coluna in ('RamalNumero', 'Duracao'):
        if coluna not in df.columns:
            raise ValueError(f'Arquivo de logins sem a coluna {coluna}.')

    df = df.copy()
    df['RamalNumero'] = df['RamalNumero'].astype(str).str.strip()
    df['duracao'] = df['Duracao'].apply(para_segundos)
    df['chamadas'] = pd.to_numeric(df.get('QTDTotalAtivoReceptivo', 0), errors='coerce').fillna(0)

    resumo = df.groupby('RamalNumero').agg(
        tempo_logado=('duracao', 'sum'),
        chamadas=('chamadas', 'sum'),
        nome_login=('UsuarioNome', 'first'),
    )
    return resumo.reset_index()


def montar_relatorio(pausas: pd.DataFrame, logins: pd.DataFrame) -> pd.DataFrame:
    df = logins.merge(pausas, on='RamalNumero', how='outer')
    df['Nome'] = df['nome_login'].fillna(df['nome_pausa']).fillna('')

    numericas = ['tempo_logado', 'chamadas', 'pausa_produtiva', 'pausa_improdutiva',
                 'qtd_produtivas', 'qtd_improdutivas', 'almoco', 'cafe', 'nao_disponivel']
    df[numericas] = df[numericas].fillna(0).astype(int)

    total_pausas = df['pausa_produtiva'] + df['pausa_improdutiva']
    disponivel = (df['tempo_logado'] - total_pausas).clip(lower=0)
    nd_por_chamada = np.where(df['chamadas'] > 0, df['nao_disponivel'] / df['chamadas'].replace(0, 1), 0)

    relatorio = pd.DataFrame({
        'Ramal': df['RamalNumero'],
        'Nome': df['Nome'],
        'Tempo logado': df['tempo_logado'].apply(hms),
        'Tempo pausa produtiva': df['pausa_produtiva'].apply(hms),
        'Qtd pausas produtivas': df['qtd_produtivas'],
        'Tempo pausa improdutiva': df['pausa_improdutiva'].apply(hms),
        'Qtd pausas improdutivas': df['qtd_improdutivas'],
        'Total de pausas': total_pausas.apply(hms),
        'Tempo disponível (sem pausa)': disponivel.apply(hms),
        'Status almoço': np.where(df['almoco'] > LIMITE_ALMOCO, 'ESTOURO', 'OK'),
        'Pausa almoço': df['almoco'].apply(hms),
        'Status café': np.where(df['cafe'] > LIMITE_CAFE, 'ESTOURO', 'OK'),
        'Pausa café': df['cafe'].apply(hms),
        'Pausa não disponível': df['nao_disponivel'].apply(hms),
        'Não disponível por chamada': pd.Series(nd_por_chamada).round().apply(hms),
    })
    return relatorio.sort_values('Ramal').reset_index(drop=True)


def painel_estouro(relatorio: pd.DataFrame) -> pd.DataFrame:
    estourou = (relatorio['Status almoço'] == 'ESTOURO') | (relatorio['Status café'] == 'ESTOURO')
    colunas = ['Ramal', 'Nome', 'Status almoço', 'Pausa almoço', 'Status café', 'Pausa café',
               'Tempo logado', 'Total de pausas', 'Tempo disponível (sem pausa)']
    return relatorio.loc[estourou, colunas].reset_index(drop=True)


def salvar_excel(relatorio: pd.DataFrame, saida: Path) -> None:
    with pd.ExcelWriter(saida, engine='openpyxl') as writer:
        relatorio.to_excel(writer, index=False, sheet_name='Resumo')
        painel_estouro(relatorio).to_excel(writer, index=False, sheet_name='Painel Estouro')

    wb = load_workbook(saida)
    vermelho = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    for ws in wb.worksheets:
        for celula in ws[1]:
            celula.font = Font(bold=True)
        for coluna in ws.iter_cols(min_row=1):
            if coluna[0].value in ('Status almoço', 'Status café'):
                for celula in coluna[1:]:
                    if celula.value == 'ESTOURO':
                        celula.fill = vermelho
            largura = max(len(str(c.value or '')) for c in coluna)
            ws.column_dimensions[coluna[0].column_letter].width = min(largura + 2, 40)
    wb.save(saida)


def main() -> None:
    parser = argparse.ArgumentParser(description='Cruza pausas e logins e gera um relatório em Excel.')
    parser.add_argument('--pausas', required=True, help='CSV ou Excel com os eventos de pausa')
    parser.add_argument('--logins', required=True, help='CSV ou Excel com o tempo logado')
    parser.add_argument('--output', default='relatorio_pausas.xlsx', help='Arquivo Excel de saída')
    args = parser.parse_args()

    pausas = resumir_pausas(ler_tabela(Path(args.pausas)))
    logins = resumir_logins(ler_tabela(Path(args.logins)))
    relatorio = montar_relatorio(pausas, logins)
    salvar_excel(relatorio, Path(args.output))

    estouros = len(painel_estouro(relatorio))
    print(f'Relatório gerado: {Path(args.output).resolve()}')
    print(f'{len(relatorio)} ramais analisados, {estouros} com estouro de almoço ou café.')


if __name__ == '__main__':
    main()
