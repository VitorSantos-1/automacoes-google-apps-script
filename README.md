# Automações em Google Apps Script (Utilitários Operacionais)

Coletânea de automações em Google Apps Script extraídas da rotina operacional de um varejo. Reúne
utilitários que eliminam tarefas manuais recorrentes em planilhas — criação de abas, limpeza
programada e validação de recebimento — demonstrando a automação de processos onde o negócio realmente
acontece: na planilha do dia a dia.

> **Nota de confidencialidade:** os dados presentes neste repositório são fictícios, gerados apenas
> para demonstração. Os dados reais da operação são confidenciais e estão protegidos — nenhum dado
> real, credencial ou informação de terceiros foi incluído aqui.

---

## Visão Geral

Boa parte do trabalho operacional do varejo passa por planilhas, e boa parte desse trabalho é
repetitivo: criar as abas do mês, limpar o que ficou velho, conferir se o recebimento veio completo.
Cada script desta coletânea automatiza uma dessas rotinas, reduzindo tempo e erro humano em tarefas
que, somadas, consomem horas.

## Contexto de Negócio

Automação não precisa de sistema grande para gerar valor: eliminar uma tarefa manual repetida toda
semana já libera tempo e reduz falha. Estes utilitários atacam gargalos concretos da operação —
padronização de abas, higiene de dados e validação de entrada — o tipo de ganho incremental que se
acumula ao longo do ano.

## Scripts

- **`criar_abas_mensais.gs`** — gera automaticamente as 12 abas mensais a partir de um modelo `PADRÃO`.
- **`limpar_abas_antigas.gs`** — apaga abas ocultas com mais de 30 dias (limpeza programada).
- **`validar_recebimento_cpd.gs`** — valida colunas obrigatórias no recebimento (CPD / NF-e).
- **`mix_oferta_delete.gs`** — rotina de manutenção do mix de ofertas.

## Impacto e Valor Gerado

- Elimina tarefas manuais recorrentes em planilhas (criação de abas, limpeza, validação).
- Reduz erro humano em rotinas operacionais.
- Padroniza processos diretamente no ambiente que a operação já usa.

## Stack

Google Apps Script - JavaScript - Google Sheets - Automação.

## Como Usar

Cole o script desejado no editor de Apps Script da planilha correspondente.

## Autor

José Vitor Santos Pinheiro — Análise de Dados e Inteligência Comercial (Varejo e Supply Chain).
Contato: vytorsantt@gmail.com
