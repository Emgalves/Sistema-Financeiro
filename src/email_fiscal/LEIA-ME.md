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

## Uso diário

1. (Se não agendou) Entrada de Dados → selecione o cliente → **📧 E-mail** → **📥 Buscar e-mails agora**.
2. A janela lista as sugestões **PENDENTES do cliente selecionado**. Linhas amarelas têm alerta: leia a coluna ALERTAS.
3. Itens sem cliente definido: abra a planilha de triagem, escolha o CLIENTE na lista suspensa, salve, feche e clique **🔄 Recarregar** (ou marque "Mostrar também itens SEM cliente definido" com o cliente certo selecionado).
4. Selecione os itens → **✅ Importar selecionados** → escolha a Etapa da Obra → confira no visualizador → **Enviar**.
5. Na próxima abertura, os itens salvos viram **IMPORTADO** sozinhos (o sistema encontra `IMPORTADO EMAIL #<id>` na OBSERVAÇÃO). A busca é pelo número exato: `#5` não é confundido com `#50` ou `#512`.

> **Feche a planilha de triagem antes de abrir a janela.** Ao abrir ou recarregar, a janela grava na `Documentos_Recebidos.xlsx` as correções automáticas descritas abaixo. Se a planilha estiver aberta no Excel, a gravação falha, as correções valem só para aquela sessão e o aviso "Não foi possível salvar..." aparece nas primeiras linhas do quadro de mensagens (role para cima para vê-lo). Para consultar a planilha, use **📊 Abrir planilha**, feche o Excel e clique **🔄 Recarregar**.

### Botões da janela "📧 E-mail"

| Botão | O que faz |
|---|---|
| ✅ Importar selecionados | Leva os itens para a lista de envio (boleto repetido na seleção é descartado) |
| 📄 Abrir documento | Abre o PDF/XML salvo do item (também com duplo clique na linha) |
| ✉️ Abrir e-mail | Abre no Gmail a mensagem recebida (para lote encaminhado, abre o lote; o remetente e o assunto originais aparecem no quadro de mensagens) |
| ✏️ Informar CNPJ | Para itens "CNPJ do fornecedor não identificado": você digita o CNPJ lido no documento |
| 🏢 Cadastrar fornecedor | Abre a tela "Novo Fornecedor" do sistema já preenchida (XML + filial já cadastrada); você completa e salva |
| 📥 Buscar e-mails agora / 📊 Abrir planilha / 🔄 Recarregar | Como antes |

A coluna ALERTAS da lista corta textos longos. Clicando em **um único item**, os alertas completos dele aparecem no quadro de mensagens, abaixo da lista. Com vários itens selecionados (como ao abrir a janela, que seleciona todos), nada é escrito.

Ao abrir/recarregar a janela, boleto e nota do mesmo fornecedor recebidos em e-mails diferentes (mesmo CNPJ, valor, vencimento e cliente) são **unidos** numa só sugestão; a do boleto fica IGNORADO com "UNIDO ao ID n". Também as sugestões repetidas (mesmo boleto, mesma parcela da mesma nota ou mesmo arquivo) viram **IGNORADO** com o alerta "REPETIDO do ID n".

### Correções automáticas ao abrir a janela

- **Releitura:** sugestões PENDENTES de boleto sem CNPJ, ou com vencimento divergente, são relidas do PDF salvo com as regras atuais (uma vez; o alerta passa a mostrar "Relido: ..."). Não é preciso coletar de novo depois de atualizar o módulo.
- **Vencimento do emissor errado:** se o vencimento do código de barras não está impresso no boleto e há uma data impressa que difere dele em múltiplo exato de 1000 dias (erro conhecido de fator de vencimento), vale a data impressa, com alerta.
- **Mesma NF do mesmo emitente:** nota sem vencimento (NF só em PDF, NFS-e, XML sem duplicatas) com NF e CNPJ iguais a um lançamento existente vira JA_LANCADO se o valor for igual ou desconhecido; com valor diferente, fica PENDENTE com "POSSÍVEL DUPLICIDADE: NF ... já lançada com outro valor".
- **Motivo da diferença de valor:** quando o alerta é "já lançada com outro valor", a janela lê as informações complementares do XML salvo da nota (NF-e: `infCpl`; NFS-e: `xInfComp` ou `OutrasInformacoes`). Se houver trechos com as palavras crédito, abater/abatimento, desconto, devolução, bonificação, adiantamento ou sinal, eles são acrescentados ao alerta depois de "Nota informa:". Exemplo real: "Nota informa: ABATER CREDITO DE R$231,63/ DIFERENCA DE R$1.630,12 BRWW BOLETO". Textos fiscais padrão (crédito de ICMS, Simples Nacional, PIS, COFINS, IPI) são descartados. Cada item é complementado uma única vez. Não funciona para nota recebida só em PDF. Se um fornecedor explicar a diferença com outras palavras, inclua-as na lista `PALAVRAS_DIFERENCA`, no `importar_para_sistema.py`.

### Versão

Todos os arquivos têm a mesma `VERSAO`. Ao abrir a janela, a primeira mensagem mostra a versão ou avisa quais arquivos estão desatualizados. A conferência compara só a linha `VERSAO`: edições que mantêm o número (como as de 09/10/2026, no histórico ao final) não são detectadas por ela.

**Depois de substituir ou editar qualquer arquivo do módulo, feche o sistema inteiro e abra de novo antes de testar.** O Python carrega cada arquivo uma única vez por execução; fechar e reabrir só a janela "📧 E-mail" continua usando a versão antiga que está na memória.

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

Para STATUS, IGNORADO e JA_LANCADO podem ser escolhidos à mão na lista suspensa. Quando a conta já existe com os mesmos dados, prefira JA_LANCADO; use IGNORADO quando você decidiu não importar por outro motivo (por exemplo, o lançamento existente já cobre a nota com outro valor).

Como a comparação com as planilhas dos clientes trata os campos:

- **CNPJ/CPF:** só os dígitos, com zeros à esquerda recompostos. Filiais da mesma empresa são reconhecidas pelos 8 primeiros dígitos.
- **NF:** zeros à esquerda e o resíduo `.0` são ignorados (`000041957`, `41957` e `41957.0` são a mesma NF). A NF escrita na REFERÊNCIA (`NF 41957`) também é reconhecida.
- **Valor:** arredondado a 2 casas, com tolerância de R$ 0,01.
- **Vencimento:** data do Excel, número serial ou texto `dd/mm/aaaa`.
- Lançamentos com STATUS diferente de ATIVO não entram na comparação.

### Alertas de POSSÍVEL DUPLICIDADE

Nos dois tipos, o item continua PENDENTE e, ao importar, o sistema pede confirmação.

| Alerta | Situação | O que fazer |
|---|---|---|
| **POSSÍVEL DUPLICIDADE com ID_LANCAMENTO n** (... confira o CNPJ) | Já existe lançamento com a mesma NF, valor e vencimento, mas no CNPJ de OUTRA empresa (8 primeiros dígitos diferentes) | Confira o CNPJ do emitente na nota. Se for a mesma conta, corrija o CNPJ do lançamento antigo, na planilha do cliente, para o CNPJ da nota: na próxima abertura da janela o item vira JA_LANCADO sozinho. Só cadastrar a filial **não** resolve, porque o lançamento antigo continua com o CNPJ errado. Se preferir não corrigir, marque IGNORADO. |
| **POSSÍVEL DUPLICIDADE: NF ... já lançada com outro valor** | Nota sem vencimento cuja NF já foi lançada pela mesma empresa com valor diferente | Leia o "Nota informa:" (se houver) e confira a nota. Se o lançamento existente já cobre a nota (crédito abatido, desconto), marque IGNORADO. Se falta uma parcela, preencha VALOR e DT_VENCTO da parcela restante e importe. |

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
- `CNPJS_EMPRESA` — CNPJs da própria empresa (hoje: RVR 30.104.762/0001-07). Nunca são tomados como fornecedor de um boleto; acrescente outros CNPJs do grupo, se houver.
- `SENHAS_PDF` — senhas para boletos em PDF protegido.
- `TOLERANCIA_VALOR` — diferença aceita ao casar boleto com parcela.

## Processar arquivos de uma pasta (WhatsApp, download manual)

```
python -m src.email_fiscal.coletor --pasta "C:\caminho\da\pasta"
```

Mesmo processamento, mesma triagem, mesma proteção contra duplicidade.

## Limitações conhecidas

- PDF escaneado (imagem) não tem texto → entra como `MANUAL` para você preencher. Abra o documento (📄), preencha na triagem `CNPJ_CPF`, `NF`, `VALOR` e `DT_VENCTO`, salve, feche o Excel e clique **🔄 Recarregar**. Se a conta já estiver lançada, o item vira JA_LANCADO sozinho; se não, preencha também `NOME` (e `TP_DESP`, se a regra de datas não se aplicar) e importe.
- NFS-e só em PDF (DANFSe do padrão nacional): CNPJ do prestador e número saem da chave de acesso de 50 dígitos; valor e vencimento, só se houver boleto junto.
- Boleto sem CNPJ impresso: o fornecedor é deduzido pela razão social cadastrada; se a empresa tiver mais de um CNPJ na base, você escolhe com "Informar CNPJ".
- Valor e vencimento vêm do código de barras. Se o valor do código não estiver impresso no boleto, aparece o alerta "ATENÇÃO: o código de barras vale R$ ..." — confirme com o fornecedor antes de pagar.
- Boleto enviado apenas como **link** (sem PDF e sem linha digitável no corpo) não é detectado.
- DANFE só em PDF: número da NF e CNPJ do emitente saem da chave de acesso; valor/vencimento, só se houver boleto junto. Peça o XML ao fornecedor.
- Guia de arrecadação/concessionária (linha começando com 8): o código não contém vencimento.
- Mantenha `Documentos_Recebidos.xlsx` **fechada** enquanto o coletor roda. Se estiver aberta, nada é gravado e os e-mails são lidos de novo na próxima execução (não há perda nem duplicidade).

---

## Histórico de correções

### 09/10/2026 (VERSAO mantida em `2026-09-26.4`)

- `roteador.py` — **marcador de importação exato.** Antes, a procura por `IMPORTADO EMAIL #5` também encontrava `#50`, `#51`, `#512`: um item pendente podia ser marcado IMPORTADO sem ter sido gravado na planilha do cliente. Corrigido antes da entrada em produção (nenhum item chegou a ser afetado).
- `roteador.py` — **NF com `.0` em texto.** A NF do e-mail e a da planilha passam pela mesma função (`_nf_texto`), que também desconsidera o resíduo `.0` gravado como texto.
- `importar_para_sistema.py` — **motivo da diferença de valor** lido do XML e acrescentado ao alerta "já lançada com outro valor" ("Nota informa: ...").
- `importar_para_sistema.py` — **alerta completo no quadro de mensagens** ao clicar em um único item.
- `importar_para_sistema.py` — **texto da confirmação de importação** corrigido: dizia sempre "mesma NF, valor e vencimento ... com CNPJ de outra empresa", o que não valia para o alerta de "outro valor".
- Fora do módulo — `limpar_nf_planilhas.py` (limpeza única, com backup em `BackUp\nf_limpeza_<data_hora>`) removeu o `.0` de 5 NFs gravadas como texto pela tela "Editar" do gerenciador de lançamentos, em 3 planilhas de clientes.
