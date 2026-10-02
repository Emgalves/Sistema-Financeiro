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
5. Na próxima abertura, os itens salvos viram **IMPORTADO** sozinhos (o sistema encontra `IMPORTADO EMAIL #<id>` na OBSERVAÇÃO).

### Significado do STATUS (planilha de triagem)

| STATUS | Significado |
|---|---|
| PENDENTE | Aguardando importação |
| IMPORTADO | Já gravado na planilha do cliente por esta rotina |
| JA_LANCADO | Já existia lançamento igual (mesmo CNPJ + valor + NF/vencimento) — não importe |
| IGNORADO | Você decidiu não importar (defina manualmente) |
| CANCELADA | Chegou por e-mail o evento de cancelamento da NF-e |

### Significado de IDENTIFICACAO

| Valor | Como o cliente foi achado | O que fazer |
|---|---|---|
| CONFIRMADO | CPF do destinatário/tomador/pagador = CPF do cliente | Nada |
| SUGERIDO | Nome completo ou CNO do cliente apareceu no texto | Conferir |
| AMBIGUO | O CPF pertence a mais de um cliente ativo | Escolher o cliente |
| NAO_IDENTIFICADO | Nada encontrado | Escolher o cliente |

---

## Ajustes (`config_email.py`)

- `TP_DESP_POR_TIPO` — **confira antes de usar.** Valores inferidos do seu histórico: NF-e e boleto → `3`, NFS-e → `5`.
- `DIAS_RETROATIVOS`, `CONSULTA_GMAIL` — período e filtro da busca (mesma sintaxe da pesquisa do Gmail; ex.: acrescentar `from:nfe@fornecedor.com.br`).
- `SENHAS_PDF` — senhas para boletos em PDF protegido.
- `TOLERANCIA_VALOR` — diferença aceita ao casar boleto com parcela.

## Processar arquivos de uma pasta (WhatsApp, download manual)

```
python -m src.email_fiscal.coletor --pasta "C:\caminho\da\pasta"
```

Mesmo processamento, mesma triagem, mesma proteção contra duplicidade.

## Limitações conhecidas

- PDF escaneado (imagem) não tem texto → entra como `MANUAL` para você preencher.
- Boleto enviado apenas como **link** (sem PDF e sem linha digitável no corpo) não é detectado.
- DANFE só em PDF: número da NF e CNPJ do emitente saem da chave de acesso; valor/vencimento, só se houver boleto junto. Peça o XML ao fornecedor.
- Guia de arrecadação/concessionária (linha começando com 8): o código não contém vencimento.
- Mantenha `Documentos_Recebidos.xlsx` **fechada** enquanto o coletor roda. Se estiver aberta, nada é gravado e os e-mails são lidos de novo na próxima execução (não há perda nem duplicidade).
