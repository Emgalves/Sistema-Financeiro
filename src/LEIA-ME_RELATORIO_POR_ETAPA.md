# Relatório por Etapa da Obra — instalação e roteiro de teste

## 1. O que há na pasta

| Arquivo | O que é | Onde colocar |
|---|---|---|
| `src/etapas_obra_core.py` | Regras, agregações, sugestões e gravação segura (sem tela) | `src/` (novo) |
| `src/relatorio_por_etapa.py` | Gera o PDF (padrão visual do relatório de despesas, logo3.png) | `src/` (novo) |
| `src/regularizacao_etapas.py` | Tela de regularização + diálogo "Ordem e marco" | `src/` (novo) |
| `patch/aplicar_patch_interface.py` | Insere o relatório na `relatorios_interface.py` | qualquer pasta |
| `src/relatorios_interface.py` | Cópia da SUA interface (a enviada nesta conversa) já com o patch | ver passo 2 |
| `patch/relatorios_interface.diff` | As 166 linhas acrescentadas (nenhuma linha existente foi alterada) | conferência |
| `tests/teste_etapas_obra.py` | Testes automáticos do núcleo, sempre em cópia da planilha | `tests/` |

## 2. Instalação (4 passos)

1. Copie os 3 arquivos novos de `src/` para a pasta `src/` do sistema. Não há dependência nova (reportlab, pandas, openpyxl e dateutil já são usados).
2. Integração na interface. Prefira rodar o patch no SEU arquivo atual, para não sobrescrever alterações feitas depois:
   `python aplicar_patch_interface.py "C:\...\src\relatorios_interface.py"`
   (cria antes `relatorios_interface.py.bak_etapas`; se qualquer âncora não for encontrada exatamente 1 vez, não altera nada). Alternativa: substituir pelo `src/relatorios_interface.py` desta pasta, se o seu ainda for o mesmo enviado.
3. `logo3.png` deve estar onde o sistema já a procura (pasta do programa). Sem ela, o PDF sai sem logo, sem erro.
4. Em Configurações, confirme que existe a etapa **CUSTOS GERAIS DA OBRA** (a comparação ignora maiúsculas/acentos). Se o texto cadastrado for outro, o painel mostra um aviso vermelho e a constante `ETAPA_GERAIS` (início de `etapas_obra_core.py`) deve ser ajustada.

Se o sistema for empacotado (PyInstaller), refaça o build; os módulos são importados dentro de funções, no mesmo padrão dos relatórios existentes.

## 3. Como o marco funciona (por cliente, por quinzena)

- O marco é uma quinzena escolhida **para cada cliente** em Relatórios > Relatório por Etapa da Obra > *Ordem e marco...*, entre as quinzenas que o próprio cliente tem. Fica salvo na aba `ETAPAS_OBRA` da planilha do cliente.
- Lançamentos **sem etapa e com DATA_REL anterior ao marco** não são pendência: ficam como estão e aparecem no bloco "Anteriores ao controle por etapa (antes de dd/mm/aaaa)". O diálogo mostra, antes de salvar, quantos lançamentos saem e entram na fila com o marco escolhido.
- Para atualizar clientes antigos aos poucos: regularize o que está na fila, recue o marco para uma quinzena anterior, e a fila passa a incluir o período recuado. Repita.
- Enquanto um cliente não tiver marco salvo, vale `MARCO_PADRAO = 05/09/2025` (constante em `etapas_obra_core.py`), a data em que o campo entrou em uso. Ela só é o ponto de partida; nada obriga a mantê-la.
- O marco só decide o que é "pendência" e o que é "anterior". Nunca altera valores nem o total da obra.

## 4. Regras aplicadas (resumo)

- TAX, ADM e TP: sempre em Custos Gerais (linha 1), mesmo com etapa gravada.
- Etapa = CUSTOS GERAIS DA OBRA: Custos Gerais, linha 2 (marcados).
- Etapa válida da lista: essa etapa.
- Sem etapa a partir do marco: Custos Gerais, linha 3, destacada como pendente de classificação.
- Sem etapa antes do marco: bloco "Anteriores".
- Mão de obra: sugerida pela etapa em andamento na quinzena (≥ 80%). Material, locação, serviço e diversos: contrato de medição, depois fornecedor já classificado (≥ 3 lançamentos, ≥ 80%), depois palavra-chave; a data não é critério. Mensalidade de motoboy (DIV): etapa em andamento na quinzena.
- Ordem das etapas: cronológica (data mediana) ou a salva por cliente. Etapa nova fora da ordem salva entra ao final.

## 5. Segurança da gravação

Só são gravadas as colunas P (HISTORICO_ALTERACAO) e Q (ETAPA_OBRA) de linhas com Q vazia, e a aba `ETAPAS_OBRA`. Antes: backup em `Backups_Sistema\<cliente>_ANTES_ETAPAS_aaaammdd_hhmmss.xlsx`. Cada linha é conferida na hora de gravar (ID, valor, status, Q vazia); a gravação vai para um arquivo temporário, é conferida e só então substitui o original. Planilha aberta no Excel ou alterada por outra pessoa no meio: nada é gravado e o original fica intacto. Nenhuma linha é inserida, apagada ou reordenada (a aba Vinculacoes depende do número da linha).

Para desfazer uma gravação: feche o Excel e copie o backup de volta sobre a planilha do cliente. Os backups não são apagados automaticamente.

## 6. Roteiro de teste sugerido

1. **Só ler:** abra o painel, escolha o cliente e confira o quadro "Situação das etapas". Clique em *Gerar Relatório*. Isso não altera a planilha. Confira o total do PDF com o total ativo do cliente.
2. **Escolha inicial:** comece pelos clientes mais novos (na cópia do clientes.xlsx enviada, por Data Inicial: HAROLDO VIEIRA DO NASCIMENTO, BRUNO AUGUSTO DELLI ZOTTI SOUZA, SÉRGIO PASSOS RIBEIRO DE CAMPOS, EDUARDO DE CARVALHO LIMA).
3. **Primeira gravação:** em *Regularizar etapas...*, aceite só as sugestões "Alta", abra a *Prévia da gravação*, confira e grave. Gere o PDF de novo: o total não muda; o valor sai de "SEM ETAPA INFORMADA" e entra nas etapas.
4. **Depois:** trate as linhas sem sugestão à mão, ajuste a ordem se quiser e, nos clientes antigos, recue o marco aos poucos.
5. Se algo estranho aparecer, guarde o backup e a mensagem exibida e me envie.

## 7. Testes executados

- `tests/teste_etapas_obra.py`: 23 verificações por planilha, todas passando nas 3 obras testadas (NOEL, GILMAR, CLEVER), sempre em cópia. Cobrem: soma dos blocos = total; gravação só em P e Q; abas, fórmulas e demais células intactas; backup; segunda gravação pulada; ID/valor/etapa inválidos; histórico com no máximo 5 entradas; arquivo bloqueado; alteração concorrente; ordem e marco.
- Tela de regularização e painel novo: exercitados em Linux com tkinter (carga, filtros, atribuição, prévia, gravação, ordem e marco, geração do PDF).
- **Não testado:** no Windows com o Excel realmente aberto (o bloqueio foi simulado), com o `GerenciadorConfiguracoes` real (usei uma lista igual à sua de 31 etapas) e com as suas outras planilhas.
