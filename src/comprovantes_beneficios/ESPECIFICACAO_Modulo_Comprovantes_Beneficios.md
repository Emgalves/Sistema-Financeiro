# Especificação — Módulo de Emissão de Comprovantes de Benefícios

**Sistema:** Vasconcelos & Rinaldi Engenharia — Sistema de Gestão Financeira
**Local no sistema:** substitui o menu "Despesas Rateadas" (não utilizado)
**Última revisão:** 12/08/2026 (v2 — remoção de Node.js/LibreOffice/SQLite, ver seção 12)

Este documento é a fonte única de verdade deste módulo. Qualquer decisão tomada em conversa que não estiver aqui deve ser considerada provisória até ser incorporada a este arquivo.

---

## 1. Objetivo

Permitir selecionar, por cliente (pagador) e competência, os colaboradores aptos a cada um dos 4 benefícios, e emitir para cada colaborador selecionado **um comprovante em PDF individual** (nunca agrupado por colaborador nem por benefício).

Benefícios:
| Benefício | Tem valor? | Campos no comprovante |
|---|---|---|
| Cesta Básica | Não | Nome, CPF, mês/ano de referência |
| Cesta de Natal | Não | Nome, CPF, ano de referência |
| Transporte | Sim | Nome, CPF, valor, valor por extenso, mês de uso, data de vencimento |
| Café | Sim | Nome, CPF, valor, valor por extenso, mês de uso, data de vencimento |

---

## 2. Regras de elegibilidade (candidatos)

**Não retroatividade:** a emissão é uma rotina para daqui em diante — competências anteriores ao mês corrente já foram resolvidas fora do sistema e não devem aparecer como opção. O seletor de competência só oferece o mês corrente e meses futuros já lançados na planilha (ex.: vencimentos antecipados). Isso não afeta o **controle** (seção 6), que continua consultável para qualquer competência já emitida por este módulo.

Fonte: aba `Dados` da planilha do cliente (ex.: `CLEVER_LUIZ_SALVADOR.xlsx`), coluna `REFERÊNCIA`.

- **Transporte**: linhas com `REFERÊNCIA == 'TRANSPORTE'` na competência selecionada.
- **Café**: linhas com `REFERÊNCIA == 'CAFÉ'` na competência selecionada.
- **Cesta Básica / Cesta de Natal**: união de CPFs únicos que aparecem em `TRANSPORTE`, `CAFÉ` ou `DIÁRIA` na competência selecionada (sem valor — é apenas confirmação de entrega física).
- Linhas com `REFERÊNCIA == 'VT E CAFÉ'` (valor combinado de Transporte+Café em uma única linha) **são ignoradas** na seleção — ficam de fora até o lançamento ser corrigido na planilha de origem, separando as duas linhas. *(decisão confirmada)*

### 2.1 Tratamento de inconsistências nos dados de origem

Identificadas na planilha real durante a análise — o módulo trata automaticamente:

- **CPF não padronizado** (`07573876670`, `003.544.326-05`, `00007573876670`): normalizado internamente para 11 dígitos numéricos, usado como chave única de identificação. A formatação exibida no PDF segue o padrão `000.000.000-00`.
- **Nome divergente para o mesmo CPF** (ex.: "ROBEVAL" vs "ROBERVAL"): resolvido usando **`Base_Fornecedores.xlsx`** (aba `Fornecedores`, filtro `tipo_pessoa == 'PF'`) como fonte canônica do nome por CPF — essa base foi conferida e não tem conflito interno (575 CPFs de pessoa física, 0 divergência). Quando o CPF não é encontrado na base de fornecedores, usa-se o nome do lançamento mais recente na planilha do cliente, com aviso na interface para cadastro posterior.

Dados do pagador (nome, endereço e cidade) são lidos de **`Clientes.xlsx`** (colunas `Nome`, `Endereço`, `Cidade`), casando pelo nome do cliente selecionado — não mais da aba `RESUMO` da planilha do cliente. **Revisão (v4):** a aba `RESUMO` está sendo descontinuada e diversos clientes reais já não a possuem (ex.: `BRUNO_AUGUSTO_DELLI_ZOTTI_SOUZA.xlsx`, só tem `Dados` e `Contratos_ADM`) — a versão anterior deste módulo dependia dela e quebrava nesses casos.

`cidade_emissao` (usada na linha de local/data do recibo) vem só da coluna **`Cidade`**, sem UF — a coluna `Endereço` já é uma concatenação de Logradouro/Número/Localidade/CEP/Estado (inclusive em cadastros antigos, com formatação inconsistente de UF: `" - MG"`, `", MG"`, `" / MG"`), então incluir UF de novo duplicaria a informação. Fallback `"NOVA LIMA"` só quando a coluna `Cidade` está vazia.

---

## 3. Formato de saída

- Comprovante final: **PDF individual**, um arquivo por colaborador por benefício.
- Fluxo de geração: **direto em Python com `reportlab`** (biblioteca pura, sem dependências externas — ver seção 12 sobre a mudança de abordagem).
- Nome de arquivo: `RECIBO_{BENEFICIO}_{CPF}_{COMPETENCIA}.pdf`
  Ex.: `RECIBO_TRANSPORTE_00354432605_2026-08.pdf`
- Local de gravação: subpasta por cliente e competência dentro da estrutura já usada por `PASTA_CLIENTES`, ex.:
  `PASTA_CLIENTES/{cliente}/Comprovantes/{competencia}/`

---

## 4. Templates de recibo

**Revisão (v3):** texto legal e layout substituídos pelo modelo definitivo `RECIBOS.docx` fornecido pelo usuário, após ver os primeiros PDFs reais gerados a partir dos modelos antigos do cliente. Mudanças em relação à v2:

- **Sem caixa/borda** — texto solto na página.
- **Título por tipo**: "RECIBO DE ENTREGA DE CESTA BASICA" / "RECIBO DE ENTREGA DE CESTA DE NATAL" / "RECIBO DE VALE TRANSPORTE" / "RECIBO DE VALE CAFÉ" (grafia "BASICA" sem acento mantida exatamente como no `RECIBOS.docx`).
- **Abertura formal, igual para os 4 tipos**: *"Pelo presente declaro para os devidos fins, que recebi de **{pagador}**, {endereço}, ..."* — nome do pagador em negrito, endereço sem negrito, com vírgula ao final.
- **Cesta Básica**: *"...a **CESTA BASICA** referente ao {mês} de {ano}."*
- **Cesta de Natal**: *"...a **CESTA DE NATAL**, como gratificação pela prestação de serviços durante o ano de {ano}."*
- **Transporte**: *"...o valor de R$ {valor} ({valor por extenso, minúsculo, sem negrito}), referente ao **VALE TRANSPORTE** para uso em {mês} de {ano}, juntamente com minha remuneração mensal."*
- **Café**: mesma estrutura, com **VALE CAFÉ** e *"a ser consumido em {mês} de {ano}"*.
- **Linha de local/data**: sem negrito, sem ponto final. Cestas: dia em branco (espaço reservado — preenchimento manual na assinatura física). Transporte/Café: dia do **vencimento real do lançamento** (`DT_VENCTO` da aba `Dados`, já usado no resto do módulo); só na ausência dele, cai no cálculo do 5º dia útil via `feriados.py` (ver seção 4.1).
- **Assinatura**: linha + nome + CPF, alinhados à **esquerda**, sem negrito, sem o rótulo "Nome:" (só o nome puro, como no `RECIBOS.docx`).

Layout implementado direto com `reportlab` (Paragraphs em sequência, sem Table/borda).

### 4.1 Dia de vencimento em Transporte/Café — `feriados.py`

O usuário forneceu `src/feriados.py`, módulo central de cálculo de dia útil (feriados nacional/estadual/municipal, sede em Belo Horizonte/MG) já usado — ou a ser usado — em outros pontos do sistema. `gerador_recibo.py` importa `calcular_enesimo_dia_util` de lá como fallback, mas a fonte primária é sempre a data de vencimento já lançada na planilha (mais confiável, pois reflete o que já foi decidido operacionalmente — um teste com dados reais mostrou 1 dia de diferença entre o cálculo puro do 5º dia útil e o vencimento realmente lançado em agosto/2026, provavelmente pela regra de sábado contar como dia útil).

Depende do pacote `holidays` (`pip install holidays`) — acrescentado aos pré-requisitos (seção 10).

Campos dinâmicos por tipo:

- **Cesta Básica**: pagador, mês/ano de referência, data de emissão (dia em branco), nome, CPF.
- **Cesta de Natal**: pagador, ano de referência, data de emissão (dia em branco), nome, CPF.
- **Transporte**: pagador, valor (numérico + por extenso), mês de uso, dia de vencimento, nome, CPF.
- **Café**: pagador, valor (numérico + por extenso), mês de uso, dia de vencimento, nome, CPF.

Valor por extenso gerado via `num2words` (pt_BR); no corpo do recibo aparece em minúsculo e sem negrito (diferente da versão v2, que era maiúsculo e em negrito).

---

## 5. Interface (dentro do menu "Despesas Rateadas" → renomeado)

Fluxo de tela:

1. **Selecionar cliente** — lista vem de `obter_clientes_ativos()` (`src/config/utils.py`), que lê `Clientes.xlsx` e considera ativo quem tem a coluna E (`Data Final`) vazia **ou** com data futura. O arquivo `.xlsx` de cada cliente em `PASTA_CLIENTES` é resolvido por correspondência tolerante a acento/espaço/underscore (o nome em `Clientes.xlsx` usa espaços, o arquivo real usa underscore — ex. `"CLEVER LUIZ SALVADOR"` → `CLEVER_LUIZ_SALVADOR.xlsx`). Cliente ativo sem arquivo correspondente aparece como aviso, não trava a tela. *(revisão — versão anterior lia direto a pasta, sem checar se o cliente estava ativo em `Clientes.xlsx`)*
2. **Selecionar competência** (mês/ano, apenas mês corrente/futuro — seção 2).
3. **Selecionar benefício** (Cesta Básica / Cesta de Natal / Transporte / Café).
4. Sistema monta a lista de candidatos elegíveis (regras da seção 2), já mostrando:
   - quem já tem comprovante emitido para essa combinação (cliente+competência+benefício) — marcado visualmente, com opção de **reemitir (2ª via)**;
   - avisos de nome divergente por CPF (seção 2.1).
5. Usuário seleciona (multi-seleção na tabela) um ou mais colaboradores e clica **Emitir selecionados**.
6. Sistema gera um PDF individual por colaborador selecionado e grava todos os registros de controle **em uma única operação de lote** (seção 6), exibindo resumo da emissão e abrindo a pasta de saída.

Botão **"Fechar"** ao lado de "Buscar candidatos" — fecha a janela sem depender do X do Windows.

Título do card no menu principal (era "Despesas Rateadas"): **"Comprovantes de Benefícios"**.

Testado de ponta a ponta com captura de tela (display virtual) e dados reais — ver seção 11.

---

## 6. Controle de emissões

**Decisão revisada (substitui SQLite — ver seção 12):** uma aba própria, **`Controle_Comprovantes`**, dentro da própria planilha do cliente (mesmo arquivo que já contém `Dados`, `RESUMO`, `Contratos_ADM` etc.) — mantém coerência com o resto do sistema, elimina a necessidade de arquivo/pasta novos, e evita depender de infraestrutura de banco de dados centralizada, que não se encaixa no modelo real de uso (cada colaborador roda o sistema na própria máquina, dados no Google Drive).

### 6.1 Por que uma aba, e por que é segura

- **Consistência**: mesmo padrão de leitura/escrita (`openpyxl`) já usado no resto do sistema para o mesmo arquivo.
- **Visibilidade automática**: por estar no Drive junto com o resto dos dados do cliente, fica visível de qualquer máquina sem sincronização extra.
- **Testado com arquivo real antes de adotar**: abrir + acrescentar aba + salvar com `openpyxl` foi comparado byte a byte (contagem de fórmulas, imagens, mesclagens) e visualmente (renderização da aba com logo e 800 fórmulas) entre o arquivo original e o resultado — nenhuma perda ou alteração detectada. Tempo de operação: < 1s no arquivo real (180 KB, 1777 linhas em `Dados`, 800 fórmulas em `RESUMO`).

### 6.2 Princípio contra corrupção por sincronização (Drive)

- **Nunca se edita uma linha já gravada.** Toda operação só **acrescenta** linhas novas ao final da aba — inclusive reemissão (2ª via) e cancelamento, que também são novas linhas, nunca uma edição da linha original. Isso preserva o histórico e minimiza a janela de conflito.
- **Emissões em lote abrem e salvam o arquivo do cliente uma única vez** (não uma vez por colaborador selecionado) — reduz tanto o tempo total quanto o número de janelas de gravação por operação.
- Resíduo de risco aceito conscientemente: como qualquer escrita em arquivo dentro do Drive, uma coincidência exata de duas gravações simultâneas (duas pessoas, duas máquinas, mesmo instante) ainda poderia colidir. Dado o padrão de uso confirmado — uma pessoa responsável, emissão mensal — esse risco é considerado desprezível; uma solução totalmente livre desse risco exigiria um serviço central dedicado, o que não se justifica para este volume de uso.

### 6.3 Estrutura da aba `Controle_Comprovantes`

Colunas (nessa ordem, cabeçalho na linha 1):

```
DATA_EMISSAO | BENEFICIO | COMPETENCIA | CPF | NOME | VALOR | DIAS |
DATA_VENCIMENTO | USUARIO | MAQUINA | CAMINHO_PDF | STATUS | OBSERVACAO
```

`STATUS`: `EMITIDO` | `CANCELADO`. Reemissão = nova linha `EMITIDO` após uma checagem que considera **a última linha gravada** para aquela combinação (benefício + competência + CPF) como o status vigente.

Verificação de duplicidade: antes de gravar, o lote inteiro é comparado contra os registros já existentes lidos em memória (uma única leitura); quem já está `EMITIDO` e não teve "permitir reemissão" marcado é pulado (não gravado de novo), e aparece no resumo final da operação.

---

## 7. Estrutura de código

```
src/
  comprovantes_beneficios/
    __init__.py
    interface.py            # janela Tkinter (fluxo da seção 5)
    dados_candidatos.py     # lê planilha do cliente + Base_Fornecedores, aplica regras da seção 2
    normalizacao.py         # normaliza CPF, nome, valor por extenso
    gerador_recibo.py       # gera o PDF diretamente com reportlab
    controle_registros.py   # lê/grava a aba Controle_Comprovantes (seção 6)
```

Integração em `sistema_principal.py`: o método `abrir_despesas_rateadas` é substituído por `abrir_comprovantes_beneficios`, que importa `InterfaceComprovantesBeneficios` e mantém o mesmo padrão de abertura (`tk.Toplevel`). Patch completo fornecido em `PATCH_sistema_principal_v2.py`.

---

## 8. Trabalho futuro (fora do escopo deste módulo)

- **Geração da planilha de "ponto diário" pelo próprio sistema**: hoje esse arquivo (ex.: `FERNANDA_-_1º_QUINZENA_JULHO2026.xlsx`, enviado à obra para registro de presença dos diaristas) é montado manualmente, o que é a origem das divergências de grafia de nome tratadas na seção 2.1. A ideia, já levantada antes e não implementada, é o sistema montar esse arquivo a partir da `Base_Fornecedores.xlsx` (seleção de cliente + colaboradores do período), eliminando a inconsistência na origem em vez de corrigi-la depois. Não bloqueia este módulo — a divergência já é resolvida aqui via consulta à base de fornecedores — mas vale ser retomado como evolução separada.

---

## 9. Modelo real de uso confirmado (importante para qualquer decisão futura de arquitetura)

Registrado aqui porque já causou um retrabalho nesta especificação (seção 12) e não deve ser reperguntado:

- Cada colaborador roda o sistema **na própria máquina**, via atalho — o `.exe` abre na tela normal do computador da pessoa, junto com Word/Excel. Não é Terminal Server/RDP.
- O `.exe` é atualizado copiando/colando manualmente (via AnyDesk) no servidor do cliente; `S:\Gestão\` é o local onde ele fica hospedado, mas a execução acontece na máquina de quem usa.
- Todos os arquivos de dados (planilhas, JSON) ficam no Google Drive, sincronizado.
- Este módulo especificamente (emissão de comprovantes) é operado por **uma pessoa responsável, uma vez por mês**, ainda que, em tese, qualquer colaborador tenha acesso a qualquer módulo.

Qualquer proposta de arquitetura (banco de dados, serviço externo, dependência de sistema) precisa ser compatível com "várias máquinas Windows independentes, dados só compartilhados via Google Drive, sem servidor de aplicação central" — não presumir Terminal Server nem servidor único de execução.

---

## 10. Pré-requisitos de implantação

Bem mais simples que a versão anterior desta especificação (seção 12):

1. **Nenhuma instalação adicional em nenhuma máquina.** `reportlab`, `num2words` e `holidays` são bibliotecas Python puras — empacotam junto do `.exe` do sistema exatamente como `openpyxl` já empacota hoje.
2. Ao gerar o `.exe` (PyInstaller ou equivalente), garantir que estejam no ambiente de build:
   ```
   pip install reportlab num2words holidays
   ```
3. `src/feriados.py` precisa existir no projeto (fornecido pelo usuário) — `gerador_recibo.py` importa `calcular_enesimo_dia_util` de lá.
4. Nenhuma constante nova em `config.py` é necessária — o módulo usa `PASTA_CLIENTES` e `ARQUIVO_FORNECEDORES`, que já existem.

---

## 11. Status de implementação

| Item | Arquivo | Status |
|---|---|---|
| 1 | `normalizacao.py` | ✅ concluído e testado |
| 2 | `dados_candidatos.py` | ✅ concluído e testado — v4: dados do pagador vêm de `Clientes.xlsx` (Nome/Endereço/Cidade), não mais da aba `RESUMO` |
| 3 | `gerador_recibo.py` | ✅ concluído e testado — `reportlab` (v2), texto/layout do `RECIBOS.docx` (v3); v5: sincronizado com ajustes finos feitos em produção (texto "referente ao **mês de** {mês}" em Cesta Básica, mais espaço entre título e corpo) |
| 4 | `controle_registros.py` | ✅ concluído e testado — reescrito com aba na planilha do cliente (v2, sem SQLite) |
| 5 | `interface.py` | ✅ concluído e testado — v4: novo parâmetro `arquivo_clientes` (necessário para o pagador vir de `Clientes.xlsx`), clientes ativos via `obter_clientes_ativos()`, botão Fechar, cabeçalho em 2 linhas |
| 6 | Integração em `sistema_principal.py` | ⏳ patch fornecido (`PATCH_sistema_principal_v2.py`), pendente de aplicação manual no arquivo real |
| 7 | Renomeação do card do menu principal | ⏳ pendente — trecho da montagem dos cards ainda não recebido (arquivo enviado até agora não incluía essa parte) |

---

## 12. Histórico de revisão de arquitetura (por que a seção 6 mudou)

Registrado para não se perder o raciocínio, já que envolveu retrabalho:

1. **Versão original**: assumi, com base numa resposta ambígua, que o sistema roda em uma única máquina/servidor (Terminal Server/RDP) acessada por todos. Com essa premissa, propus SQLite com o arquivo em `S:\Gestão\dados_locais\`, e geração de PDF via `docx-js` (Node.js) + LibreOffice, ambos instalados uma única vez "no servidor".
2. **Contradição identificada pelo usuário**: ele havia recebido orientação anterior (não desta conversa) de instalar Node.js em cada máquina de usuário — o oposto do que eu disse. Isso expôs que a premissa de "servidor único" estava errada.
3. **Modelo real confirmado** (seção 9): cada colaborador roda o sistema localmente, na própria máquina; `S:\Gestão\` é só onde o `.exe` fica hospedado/distribuído, não onde ele executa.
4. **Consequência 1 — geração de PDF**: exigir Node.js + LibreOffice em cada máquina de colaborador é inviável operacionalmente. Trocado por `reportlab` (Python puro, empacota no `.exe`).
5. **Consequência 2 — controle de emissões**: `S:\Gestão\` é unidade de rede mapeada em várias máquinas, não disco local de um servidor único — o mesmo risco de corrupção do Google Drive se aplica a SQLite ali. Descartado.
6. **Solução adotada**: o próprio usuário sugeriu usar uma aba na planilha do cliente, já que o arquivo já concentra tudo daquele cliente (Dados, contratos, medições etc.) — testado e confirmado seguro antes de adotar (seção 6.1).
