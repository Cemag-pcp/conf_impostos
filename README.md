# verificacao_nf

Automação para entrar no Innovaro ERP, baixar relatórios e consolidar/processar
os dados para verificação de notas fiscais.

## O que o script faz

1. **Etapa 1**: abre "TI-Relatório de Baixas Recursos com Chave NFe e Impostos"
   (Venda > Consultas), filtra por Emissão (Início/Fim) = hoje, coleta a tabela.
2. **Etapa 2**: abre "99003 Download de XML Manifestados (C)" (Fiscal e
   Regulamentação > Consultas > Auxiliares Fiscais > Manifestação (C)), mesmo
   período, baixa o .zip com os XMLs das NFes manifestadas.
3. **Etapa 3**: descompacta o .zip, lê cada XML e monta a planilha final
   (`downloads/verificacao_nf_<timestamp>.xlsx`) com 4 abas:
   - **Baixas** — dados da Etapa 1
   - **Resumo NFe** — um resumo por nota, extraído dos XMLs
   - **Itens** — um detalhamento por produto (só consulta)
   - **Cruzamento** — fórmulas do Excel (SUMIF/COUNTIF) comparando, por NFe,
     os impostos (ICMS/IPI/PIS/COFINS) e o total lançados em "Baixas" contra
     os valores do XML. Status por nota: `SEM BAIXA`, `OK` ou `DIVERGENTE`.
     As fórmulas usam referência de coluna inteira — arrastar a última linha
     pra baixo continua funcionando se entrarem mais NFes depois.
4. **Google Sheets**: acrescenta (não sobrescreve) as linhas de "Baixas",
   "Resumo NFe" e "Itens" desta execução na planilha Google configurada em
   `GOOGLE_SHEETS_ID` (`.env`) — histórico acumulado dia a dia, já que o
   script roda 1x/dia com período de só hoje. Precisa que a service account
   (`GOOGLE_CLIENT_EMAIL` no `.env`) já tenha sido compartilhada como editora
   na planilha, e que as 3 abas existam com esses nomes exatos. Se
   `GOOGLE_SHEETS_ID` não estiver configurado, essa etapa é pulada. Falha no
   envio ao Sheets não derruba a execução — a planilha local já foi salva.

## Uso

```
setup_projeto.bat
python main.py
```

**Roda desacompanhado (agendado de madrugada) — sempre usa a data de hoje**
como Emissão Início/Fim, sem precisar de nenhum argumento.
