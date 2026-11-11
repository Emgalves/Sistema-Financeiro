# Módulo email_fiscal — NF-e, NFS-e e boletos do Gmail para o sistema

## O que ele faz

```
Gmail ──► coletor.py ──► Documentos_Recebidos.xlsx ──► botão "📧 E-mail" ──► visualizador ──► ENVIAR ──► <CLIENTE>.xlsx
          (lê anexos)     (triagem: você confere)      (Entrada de Dados)    (o de sempre)
```

1. Procura no Gmail e-mails com anexo XML, PDF ou ZIP (somente leitura: não apaga, não move, não marca como lido).
2. Reconhece: **NF-e** (XML), **NFS-e** (XML padrão nacional e ABRASF), **evento de cancelamento de NF-e**, **boleto** (PDF ou linha digitável no corpo do e-mail), **DANFE em PDF** sem XML.
3. Descobre o **cliente** pelo CPF do destinatário/tomador/pagador comparado com a coluna CPF de `Clientes.xlsx` (e, na falta, pelo nome do cliente ou CNO no texto).
4. Descobre o **fornecedor** pelo CNPJ na `Base_Fornecedores.xlsx` (NOME e CATEGORIA vêm de lá).
5. Gera **uma sugestão por conta a pagar** (cada parcela/duplicata da NF é uma linha) e casa o boleto com a parcela certa pelo valor e vencimento — inclusive quando o boleto chega em outro e-mail depois da nota.
6. Nada vai direto para a planilha do cliente. Você importa pelo botão **📧 E-mail** e grava com **Enviar**, passando por todas as validações que o sistema já tem.

Arquivos baixados e triagem ficam em `<Financeiro>\Email_Fiscal\`, ao lado de `Planilhas_Base` e `Clientes`, **no ambiente em uso** (segue o `src/config/config.py`):

| Ambiente | Triagem e anexos | Credenciais do Gmail |
|---|---|---|
| TESTE | `C:\Users\Obras\sistema_gestao_testes\testes\Financeiro\Email_Fiscal\` | `%APPDATA%\SistemaFinanceiro\gmail\teste\` |
| PRODUÇÃO | `<BASE_DADOS>\Email_Fiscal\` (a mesma pasta Financeiro do `config_caminhos.json`) | `%APPDATA%\SistemaFinanceiro\gmail\producao\` |

Se o `config.py` cair no fallback de emergência (caminho de produção não encontrado → TESTE), o módulo acompanha e também passa a usar as pastas de TESTE.

---

## Instalação (uma vez só)

### Passo 1 — Copiar a pasta

Copie a pasta `email_fiscal` para dentro de `src`:

```
Sistema-Financeiro\src\email_fiscal\
```

### Passo 2 — Instalar bibliotecas

No terminal, na pasta `Sistema-Financeiro`:

```
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pdfplumber
```

(`openpyxl` e `python-dateutil` você já usa.)

### Passo 3 — Criar a credencial do Google (≈10 minutos)

1. Acesse https://console.cloud.google.com/ com a conta Gmail que recebe as notas.
2. No topo, **Selecionar projeto → Novo projeto** → nome `Sistema Financeiro` → Criar. Deixe-o selecionado.
3. Menu ☰ → **APIs e serviços → Biblioteca** → pesquise **Gmail API** → **Ativar**.
4. Menu ☰ → **APIs e serviços → Tela de permissão OAuth** (em contas novas aparece como **Google Auth Platform**):
   - **Branding / Informações do app**: nome `Sistema Financeiro`, seu e-mail de suporte e de contato → Salvar.
   - **Público-alvo / Audience**: tipo **Externo** (conta @gmail.com) ou **Interno** (se for Google Workspace da empresa).
   - Se for **Externo**: em *Usuários de teste*, adicione o seu próprio e-mail.
5. Menu ☰ → **APIs e serviços → Credenciais** (ou **Clients**) → **Criar credenciais → ID do cliente OAuth** → Tipo: **App para computador (Desktop app)** → Criar → **Fazer download do JSON**.
6. Renomeie o arquivo baixado para `credentials.json` e coloque em:

   ```
   TESTE   : C:\Users\<seu usuário>\AppData\Roaming\SistemaFinanceiro\gmail\teste\credentials.json
   PRODUÇÃO: C:\Users\<seu usuário>\AppData\Roaming\SistemaFinanceiro\gmail\producao\credentials.json
   ```

   (Cole `%APPDATA%` na barra do Explorer para chegar em `AppData\Roaming`; crie as pastas.) O mesmo `credentials.json` serve para os dois ambientes; o que muda é o `token.json`, criado no primeiro login de cada ambiente com a conta Gmail daquele ambiente.
   Essa pasta fica **fora** do projeto de propósito: a credencial nunca vai para o GitHub.

> **Importante — validade do acesso.** Com o app em modo **Teste** e tipo **Externo**, o Google invalida a autorização a cada **7 dias** e você teria de logar de novo toda semana. Para evitar: em **Público-alvo / Audience**, clique em **Publicar app** (*In production*). Não precisa de verificação do Google para uso próprio; no login aparecerá "o Google não verificou este app" → clique em **Avançado → Acessar Sistema Financeiro (não seguro)**. É o seu próprio app acessando a sua própria conta. Com tipo **Interno** (Workspace) isso não acontece.

### Passo 4 — Primeiro login

Na pasta `Sistema-Financeiro`:

```
python -m src.email_fiscal.coletor --dias 60
```

O navegador abre, você escolhe a conta e autoriza **"Ver suas mensagens e configurações de e-mail"** (somente leitura). O token fica salvo em `...\SistemaFinanceiro\gmail\<ambiente>\token.json` e as próximas execuções não pedem login.

Ao final, abra `<Financeiro>\Email_Fiscal\Documentos_Recebidos.xlsx` e confira a aba **Lancamentos**.

### Passo 5 — Botão no sistema

Em `src/Sistema_Entrada_Dados.py`, dentro de `setup_aba_fornecedor`, logo depois do botão **"📋 Diárias"** (procure `command=self.importar_diarias,`), acrescente:

```python
        ttk.Button(
            frame_botoes_import,
            text="📧 E-mail",
            command=self.importar_documentos_email,
            style='Medium.TButton'
        ).pack(side='left', padx=5)
```

E, dentro da classe `SistemaEntradaDados` (por exemplo logo antes de `def importar_folha_rh(self):`), acrescente o método:

```python
    def importar_documentos_email(self):
        """Importa NF-e/NFS-e/boletos lidos do Gmail (triagem em Documentos_Recebidos.xlsx)"""
        try:
            from src.email_fiscal.importar_para_sistema import abrir_importacao_email
        except ImportError:
            from email_fiscal.importar_para_sistema import abrir_importacao_email
        abrir_importacao_email(self)
```

### Passo 6 (opcional) — Buscar automaticamente de hora em hora

Use o `coletar_email.bat` (ajuste o caminho dentro dele) no **Agendador de Tarefas do Windows**:
Criar Tarefa Básica → Diariamente → Ação: Iniciar programa → selecione o `.bat` → depois, em Propriedades → Disparadores → Editar → marque **Repetir a tarefa a cada 1 hora**.

---

### Passo 7 — Botão "🔍 Documentos" nas telas de lançamentos

Em `src/Sistema_Entrada_Dados.py` há duas telas: **Gerenciar Lançamentos** (classe `GerenciadorLancamentos`) e **Lançamentos - <fornecedor>** (classe `VisualizadorLancamentosFornecedor`). Nas DUAS, logo depois do botão **"Ver Histórico"**, acrescente o botão. O frame é diferente em cada uma:

Em `GerenciadorLancamentos` (procure `text="Ver Histórico",` dentro de `frame_acoes`):

```python
        ttk.Button(frame_acoes, text="🔍 Documentos",
                command=self.ver_documentos_lancamento).pack(side='left', padx=5)
```

Em `VisualizadorLancamentosFornecedor` (procure `self.btn_ver_historico.pack(`):

```python
        ttk.Button(frame_botoes, text="🔍 Documentos",
                  command=self.ver_documentos_lancamento).pack(side='left', padx=5)
```

E acrescente este método nas DUAS classes (por exemplo, logo antes de `def visualizar_historico_lancamento(self):`):

```python
    def ver_documentos_lancamento(self):
        """Boleto, NF e e-mail de origem do lançamento selecionado"""
        try:
            from src.email_fiscal.localizar_documentos import abrir_documentos_do_lancamento
        except ImportError:
            from email_fiscal.localizar_documentos import abrir_documentos_do_lancamento
        abrir_documentos_do_lancamento(self)
```

Uso: selecione um lançamento na lista → **🔍 Documentos**. A janela mostra o que foi encontrado:

| Nível | Quando |
|---|---|
| EXATO | O lançamento foi importado pelo e-mail (OBSERVAÇÃO com "IMPORTADO EMAIL #n") ou a triagem o reconheceu como já lançado (JA_LANCADO apontando para este ID_LANCAMENTO) |
| PROVÁVEL | Documento lido da mesma empresa (raiz do CNPJ), mesmo valor e mesma NF ou mesmo vencimento |

Botões: **📄 Abrir documento** (ou duplo clique), **✉️ Abrir e-mail**, **🔎 Pesquisar no Gmail** (abre no navegador a pesquisa com a NF e o valor, no período de 120 dias antes a 15 dias depois do vencimento — útil quando nada foi encontrado no arquivo local).

## Uso diário

1. (Se não agendou) Entrada de Dados → selecione o cliente → **📧 E-mail** → **📥 Buscar e-mails agora**.
2. A janela lista as sugestões **PENDENTES do cliente selecionado**. Linhas amarelas têm alerta: leia a coluna ALERTAS.
3. Itens sem cliente definido: abra a planilha de triagem, escolha o CLIENTE na lista suspensa, salve, feche e clique **🔄 Recarregar** (ou marque "Mostrar também itens SEM cliente definido" com o cliente certo selecionado).
4. Selecione os itens → **✅ Importar selecionados** → escolha a Etapa da Obra → confira no visualizador → **Enviar**.
5. Na próxima abertura, os itens salvos viram **IMPORTADO** sozinhos (o sistema encontra `IMPORTADO EMAIL #<id>` na OBSERVAÇÃO).

### Botões da janela "📧 E-mail"

| Botão | O que faz |
|---|---|
| ✅ Importar selecionados | Leva os itens para a lista de envio (boleto repetido na seleção é descartado) |
| 📄 Abrir documento | Abre o PDF/XML salvo do item (também com duplo clique na linha) |
| ✉️ Abrir e-mail | Abre no Gmail a mensagem recebida (para lote encaminhado, abre o lote; o remetente e o assunto originais aparecem no quadro de mensagens) |
| ✏️ Informar CNPJ | Para itens "CNPJ do fornecedor não identificado": você digita o CNPJ lido no documento |
| 🏢 Cadastrar fornecedor | Abre a tela "Novo Fornecedor" do sistema já preenchida (XML + filial já cadastrada); você completa e salva |
| 🔄 Reprocessar triagem | Relê, em segundo plano, os PDFs da triagem: boletos sem CNPJ, PDFs protegidos e clientes a confirmar pelo pagador. A tela continua livre. Também roda sozinho no fim de cada coleta. |
| 📂 Ler pasta(s) | Lê XML/PDF/ZIP/EML de uma ou mais pastas (após escolher cada pasta, o sistema pergunta se quer adicionar outra), com subpastas, em segundo plano (WhatsApp, downloads, ou a pasta Documentos_Email copiada de outro ambiente). Arquivos que já estão em Documentos_Email não são copiados de novo; conteúdo idêntico é lido uma vez. Sem o e-mail, a "data do e-mail" passa a ser a data do arquivo. |
| 📥 Buscar e-mails agora / 📊 Abrir planilha / 🔄 Recarregar | Como antes |

### Triagem dentro da janela (sem abrir o Excel)

- **Mostrar:** filtra a lista — todos os pendentes, só possíveis duplicidades, só com alertas, só sem alertas (prontos) ou ignorados.
- **Possíveis duplicidades** aparecem em **vermelho** e **não vêm selecionadas**: só são importadas se você as selecionar.
- **⚖️ Comparar com o lançamento existente:** mostra lado a lado a sugestão e o lançamento da planilha apontado no alerta (diferenças em amarelo), com os botões "É duplicado: ignorar" e "Não é duplicado: manter pendente" (o alerta não volta).
- **📝 Editar sugestão:** janela para corrigir CNPJ/CPF (ao sair do campo, nome e categoria vêm da base), nome, referência, NF, valor, vencimento, categoria, TP_DESP, dados bancários e cliente, com botão para abrir o documento. Os alertas resolvidos pela edição somem. TP_DESP preenchido na edição prevalece sobre a regra de datas.
- **🚫 Ignorar selecionados:** marca IGNORADO de uma vez quantos itens estiverem selecionados.
- **↩️ Voltar para pendente:** no filtro "Ignorados", desfaz o ignorar.

Ao abrir/recarregar a janela, boleto e nota recebidos em e-mails diferentes são **unidos** numa só sugestão (a da nota recebe linha digitável, vencimento e PDF; a do boleto fica IGNORADO com "UNIDO ao ID n"):
- **mesma empresa** (raiz do CNPJ, qualquer filial), valor, vencimento (ou nota sem vencimento) e cliente, com um único candidato; ou
- boleto **sem CNPJ** ou com **CNPJ diferente** do da nota: mesmo cliente, valor exato e mesmo vencimento, havendo exatamente uma nota e um boleto nessas condições (com alerta para conferir).

**NF impressa no boleto:** o "Nº do documento" / "Nota Fiscal" do boleto (ex.: 000000018337A, 566417-3, 34824-1/1) é guardado na coluna NF_PROVAVEL e aparece na janela como "18337?". Ele é usado para:
- reconhecer como JA_LANCADO um boleto cuja NF já foi lançada pela mesma empresa e com o mesmo valor, **mesmo com vencimento diferente** (boleto reemitido, ou lançamento registrado com a data do pagamento);
- unir boleto e nota da mesma NF mesmo com vencimentos diferentes (vale o do boleto, com alerta).
Nota de parcela única com mesma NF e valor de um único lançamento da mesma empresa também é reconhecida, mesmo com outra data.

Se o boleto já foi **importado** como "BOLETO - AGUARDANDO NF" e a nota chega depois, a nota vira JA_LANCADO com o alerta "lançado SEM NF ... atualize NF e referência nesse lançamento". Também as sugestões repetidas (mesmo boleto, mesma parcela da mesma nota ou mesmo arquivo) viram **IGNORADO** com o alerta "REPETIDO do ID n".

### Correções automáticas ao abrir a janela (rápidas: sem abrir PDFs)

- **Cliente pelo cadastro atual:** sugestões PENDENTES com cliente SUGERIDO, NAO_IDENTIFICADO ou AMBIGUO são reavaliadas pelo CPF/CNPJ do destinatário (XML) ou do pagador (boleto). Se o documento já estiver na `Clientes.xlsx` (por exemplo, cadastrado depois da coleta), viram CONFIRMADO. Clientes escolhidos manualmente não mudam.
- **Releitura:** sugestões PENDENTES de boleto sem CNPJ, ou com vencimento divergente, são relidas do PDF salvo com as regras atuais (uma vez; o alerta passa a mostrar "Relido: ..."). Não é preciso coletar de novo depois de atualizar o módulo.
- **Vencimento do emissor errado:** se o vencimento do código de barras não está impresso no boleto e há uma data impressa que difere dele em múltiplo exato de 1000 dias (erro conhecido de fator de vencimento), vale a data impressa, com alerta.
- **Mesma NF do mesmo emitente:** nota sem vencimento (NF só em PDF, NFS-e, XML sem duplicatas) com NF e CNPJ iguais a um lançamento existente vira JA_LANCADO se o valor for igual ou desconhecido; com valor diferente, fica PENDENTE com "POSSÍVEL DUPLICIDADE: NF ... já lançada com outro valor".

### Visualização das notas

"Abrir documento" (na janela de importação e no 🔍 Documentos) mostra XML de NF-e/NFS-e como uma página legível no navegador: emitente, destinatário, itens, totais e parcelas. A página tem o link "ver o XML original". PDFs continuam abrindo no leitor de PDF.

### Outras correções automáticas ao abrir a janela

- Fornecedor cadastrado depois da coleta: o alerta "NÃO cadastrado"/"FILIAL NOVA" some e NOME/CATEGORIA passam a ser os do cadastro.
- NFC-e apenas **citada** nas informações complementares de outra nota (ex.: "NFCe Ref.: 3126...") não vira sugestão; as criadas por versões anteriores ficam IGNORADO.

### Ao importar

- Boleto sem NF, mas com o número lido no próprio boleto, vai com a REFERÊNCIA **"BOLETO - NF 35599 (A CONFIRMAR)"** (`NF_PROVAVEL_NA_REFERENCIA`). Quando a nota chegar, ela é reconhecida como já lançada por esse número.
- Se, depois de importar, ignorar ou decidir uma duplicidade, não sobrar nenhuma sugestão pendente do cliente, a janela fecha sozinha (`FECHAR_JANELA_QUANDO_VAZIA`). Ao abrir uma janela já vazia, ela continua aberta (para buscar e-mails ou reprocessar).

### Notas de serviço em PDF e documentos complementares

- **DANFSe (NFS-e padrão nacional) em PDF:** CNPJ e número vêm da chave de acesso; o **valor líquido** é lido pela posição na página e o **vencimento**, quando escrito nas informações complementares ("VENCIMENTO: dd/mm/aaaa"). Ambos com alerta para conferir.
- **PDF de NFS-e sem CNPJ** (ex.: relatório de medição): o fornecedor é procurado pela razão social cadastrada no texto, e o número ao lado de "NFS-e:" vira NF provável.
- **Documento complementar:** um PDF desses que cita a NF de outra sugestão da mesma empresa e cliente é anexado a ela e fica IGNORADO ("Documento complementar da NF ..."). Se a NF já estiver lançada, vira JA_LANCADO.
- Na edição, com o CNPJ vazio, digitar o **nome** do fornecedor e sair do campo procura na base (um resultado preenche o CNPJ; vários são listados).

### Notas de débito e fornecedor pelo remetente

- **Nota de débito** (ex.: locação de bens móveis, que dispensa NFS-e): quando o PDF traz a linha do quadro de cobrança "número, valor, número, vencimento" (ex.: "071845 503,29 071845 07/09/2026"), número, valor, vencimento e referência ("LOCAÇÃO - CONTRATO ...") são preenchidos, com alerta para conferir. Nesses formulários o CNPJ costuma estar na imagem de fundo e não pode ser lido.
- **CNPJ pelo remetente:** se o documento não traz CNPJ legível, o sistema usa o CNPJ das outras sugestões recebidas do mesmo domínio de e-mail (ex.: @fornecedor.com.br), desde que o domínio leve a um único CNPJ. Provedores genéricos (gmail, hotmail...) são ignorados. Não se aplica a arquivos lidos por "Ler pasta(s)", que não têm remetente.

### NF-e sem cobrança

NF-e de **remessa** ou **retorno** (natureza da operação) sem duplicatas, ou com meio de pagamento **90 - Sem pagamento**, não é conta a pagar: fica IGNORADO automaticamente, com o alerta "NF-e SEM COBRANÇA (...)". Nota com duplicatas ou boleto nunca é ignorada por essa regra. As palavras e códigos ficam em `NATUREZAS_SEM_COBRANCA` e `PAGAMENTOS_SEM_COBRANCA`; um item ignorado por engano volta pelo filtro "Ignorados (para restaurar)".

### Versão

Todos os arquivos têm a mesma `VERSAO`. Ao abrir a janela, a primeira mensagem mostra a versão ou avisa quais arquivos estão desatualizados.

### E-mails encaminhados

Funcionam os dois jeitos: **Encaminhar** normal e **Encaminhar como anexo** (inclusive vários e-mails de uma vez). No segundo caso, cada e-mail original é tratado separadamente, com o seu remetente, assunto e data. Arquivos `.eml` também são lidos (no Gmail e com `--pasta`).

### Filiais

Se o CNPJ não está na base mas outra filial da mesma empresa está (mesmos 8 primeiros dígitos), o alerta é **FILIAL NOVA** e NOME/CATEGORIA vêm do cadastro existente. Ainda assim cadastre a filial (botão 🏢), porque o pagamento vai para o CNPJ da nota/boleto.

### Significado do STATUS (planilha de triagem)

| STATUS | Significado |
|---|---|
| PENDENTE | Aguardando importação |
| IMPORTADO | Já gravado na planilha do cliente por esta rotina |
| JA_LANCADO | Já existe lançamento da MESMA EMPRESA (mesmo CNPJ ou outra filial: 8 primeiros dígitos iguais) com mesmo valor, NF e vencimento — não importe |
| IGNORADO | Você decidiu não importar (defina manualmente) |
| CANCELADA | Chegou por e-mail o evento de cancelamento da NF-e |

Alerta **POSSÍVEL DUPLICIDADE com ID_LANCAMENTO n**: já existe lançamento com a mesma NF, valor e vencimento, mas no CNPJ de OUTRA empresa. O item continua PENDENTE; ao importar, o sistema pede confirmação. Se for mesmo duplicado, marque IGNORADO na triagem (e, se for o caso, corrija o CNPJ do lançamento antigo).

### Significado de IDENTIFICACAO

| Valor | Como o cliente foi achado | O que fazer |
|---|---|---|
| CONFIRMADO | CPF do destinatário/tomador/pagador = CPF do cliente | Nada |
| SUGERIDO | Nome completo ou CNO do cliente apareceu no texto | Conferir |
| AMBIGUO | O CPF pertence a mais de um cliente ativo | Escolher o cliente |
| NAO_IDENTIFICADO | Nada encontrado | Escolher o cliente |

---

## Ajustes (`config_email.py`)

- `REGRA_TP_DESP_POR_DATA` — TP_DESP é calculado **na importação**, com a DATA_REL em uso: e-mail anterior à DATA_REL e vencimento posterior → `3`; e-mail anterior e vencimento até a DATA_REL (inclusive) → `5`. Fora disso, vale o TP_DESP digitado na triagem; vazio bloqueia a importação do item. `TP_DESP_POR_TIPO` só é usado se a regra for desligada.
- `DIAS_RETROATIVOS`, `CONSULTA_GMAIL` — período e filtro da busca (mesma sintaxe da pesquisa do Gmail; ex.: acrescentar `from:nfe@fornecedor.com.br`).
- `DATA_REL_PELO_VENCIMENTO` / `DIAS_RELATORIO` — na importação, cada item vai para o relatório (dias 5 e 20) mais recente até o seu vencimento, nunca anterior ao relatório em uso. Ex.: relatório em uso 05/11 → venc. 17/11 fica em 05/11; 23/11 e 26/11 vão para 20/11; 11/12 e 15/12 vão para 05/12. O TP_DESP é calculado com a DATA_REL do próprio item.
- `SENHAS_PELO_CPF_DOS_CLIENTES` — PDFs protegidos são abertos tentando, para cada cliente, os 3, 4, 5 e 6 primeiros dígitos do CPF/CNPJ e o documento completo, além de `SENHAS_PDF`. Itens MANUAL antigos de PDF protegido são relidos automaticamente ao abrir a janela.
- `CNPJS_EMPRESA` — CNPJs da própria empresa (hoje: RVR 30.104.762/0001-07). Nunca são tomados como fornecedor de um boleto; acrescente outros CNPJs do grupo, se houver.
- `TIMEOUT_GMAIL`, `MAX_TAMANHO_PDF_MB`, `MAX_PAGINAS_PDF`, `SALVAR_A_CADA` — proteções para coletas longas: uma mensagem que não responde em `TIMEOUT_GMAIL` segundos é pulada (e tentada de novo na próxima execução); PDFs muito grandes vão para triagem manual; só as primeiras páginas de cada PDF são lidas; a triagem é gravada a cada `SALVAR_A_CADA` mensagens.
- `SENHAS_PDF` — senhas para boletos em PDF protegido.
- `TOLERANCIA_VALOR` — diferença aceita ao casar boleto com parcela.

## Processar arquivos de uma pasta (WhatsApp, download manual)

```
python -m src.email_fiscal.coletor --pasta "C:\caminho\da\pasta"
```

Mesmo processamento, mesma triagem, mesma proteção contra duplicidade.

## Limitações conhecidas

- PDF escaneado (imagem) não tem texto → entra como `MANUAL` para você preencher.
- NFS-e só em PDF (DANFSe do padrão nacional): CNPJ do prestador e número saem da chave de acesso de 50 dígitos; valor e vencimento, só se houver boleto junto.
- Boleto sem CNPJ impresso: o fornecedor é deduzido pela razão social cadastrada; se a empresa tiver mais de um CNPJ na base, você escolhe com "Informar CNPJ".
- Valor e vencimento vêm do código de barras. Se o valor do código não estiver impresso no boleto, aparece o alerta "ATENÇÃO: o código de barras vale R$ ..." — confirme com o fornecedor antes de pagar.
- Boleto enviado apenas como **link** (sem PDF e sem linha digitável no corpo) não é detectado.
- DANFE só em PDF: número da NF e CNPJ do emitente saem da chave de acesso; valor/vencimento, só se houver boleto junto. Peça o XML ao fornecedor.
- Guia de arrecadação/concessionária (linha começando com 8): o código não contém vencimento.
- Mantenha `Documentos_Recebidos.xlsx` **fechada** enquanto o coletor roda. Se estiver aberta, nada é gravado e os e-mails são lidos de novo na próxima execução (não há perda nem duplicidade).
