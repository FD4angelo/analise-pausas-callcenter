"""Gera arquivos de exemplo com operadores fictícios, no mesmo formato da exportação do discador."""

import csv
import random
from pathlib import Path

random.seed(42)
PASTA = Path(__file__).parent / 'dados_exemplo'
PASTA.mkdir(exist_ok=True)

NOMES = ['Ana Souza', 'Bruno Lima', 'Carla Mendes', 'Diego Rocha', 'Elisa Prado',
         'Felipe Torres', 'Gabriela Nunes', 'Hugo Martins', 'Isabela Costa', 'João Ribeiro']
DATA = '03/11/2026 00:00:00'

with open(PASTA / 'tempo_logado.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f, delimiter=';')
    w.writerow(['Data', 'GrupoNome', 'RamalNumero', 'UsuarioNome', 'Duracao', 'QTDTotalAtivoReceptivo'])
    for i, nome in enumerate(NOMES):
        w.writerow([DATA, 'Equipe Exemplo', 2000 + i, nome,
                    random.randint(7 * 3600, 8 * 3600), random.randint(90, 200)])

STATUS = {
    'Almoço': (48 * 60, 66 * 60, 1),
    'Café': (12 * 60, 23 * 60, 1),
    'Banheiro': (2 * 60, 8 * 60, 2),
    'Não Disponível': (60, 10 * 60, 3),
    'Feedback': (5 * 60, 15 * 60, 1),
    'Retorno Cliente': (2 * 60, 6 * 60, 3),
    'Atendimento WhatsApp': (3 * 60, 10 * 60, 2),
}

with open(PASTA / 'tempo_pausa.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f, delimiter=';')
    w.writerow(['Data', 'RamalNumero', 'Usuario', 'RamalStatusDetalhe', 'DuracaoEvento'])
    for i, nome in enumerate(NOMES):
        for status, (minimo, maximo, vezes) in STATUS.items():
            for _ in range(random.randint(0, vezes) if status not in ('Almoço', 'Café') else 1):
                w.writerow([DATA, 2000 + i, nome, status, random.randint(minimo, maximo)])

print(f'Arquivos criados em {PASTA}')
