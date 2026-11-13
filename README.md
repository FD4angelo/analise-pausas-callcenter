# Análise de pausas de operadores

Script em Python que uso para acompanhar as pausas dos operadores de um call center. Ele junta duas exportações do discador, uma com os eventos de pausa e outra com o tempo logado, e gera uma planilha com:

- **Resumo:** por ramal, mostra tempo logado, pausas produtivas e improdutivas, tempo disponível e o tempo em "não disponível" por chamada atendida
- **Painel Estouro:** só quem passou do limite de almoço (1h) ou de café (20 min), com o status marcado em vermelho

Antes eu fazia esse cruzamento na mão no Excel, todo dia. Hoje é um comando.

Os dados da pasta `dados_exemplo` são inventados, gerados pelo `gerar_dados_exemplo.py`.

## Como rodar

```bash
pip install -r requirements.txt
python analise_pausas.py --pausas dados_exemplo/tempo_pausa.csv --logins dados_exemplo/tempo_logado.csv
```

O resultado sai em `relatorio_pausas.xlsx`. Os arquivos podem ser CSV (separado por `;`) ou Excel.

## Colunas esperadas

| Arquivo | Colunas |
|---|---|
| Pausas | `RamalNumero`, `Usuario`, `RamalStatusDetalhe`, `DuracaoEvento` (segundos) |
| Logins | `RamalNumero`, `UsuarioNome`, `Duracao` (segundos), `QTDTotalAtivoReceptivo` |

Os nomes de pausa são comparados sem acento e sem diferença de maiúsculas, então "Almoço" e "almoco" contam como a mesma pausa. As listas de pausas produtivas e improdutivas e os limites ficam no começo do script.

---

**English:** Python script that merges two call center dialer exports (pause events and login time) and builds an Excel report per agent, with a panel highlighting who went over the lunch (1h) or coffee (20 min) break limits. Sample data is fake.
