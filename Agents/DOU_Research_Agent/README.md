# Agente de Pesquisa no Diário Oficial da União (DOU)

Script em Python (`agent.py`) que atua como jornalista isento e técnico: pesquisa
o Diário Oficial da União na **fonte oficial** e gera um relatório em Word.

## 1. Visão geral

| Campo | Valor |
| --- | --- |
| Arquivo | `D:\opencode_desktop\Agents\DOU_Research_Agent\agent.py` |
| Fonte oficial | `https://pesquisa.in.gov.br/imprensa/core/jornalList.action` |
| Órgão filtrado | Ministério da Justiça e Segurança Pública (MJSP) |
| Saída | `D:\opencode_desktop\resultado\Relatorio_DOU_MJSP_<AAAAMMDD>.docx` |
| Dependências | `requests`, `pymupdf`, `python-docx` |

O script executa o fluxo:
1. Cria uma sessão HTTP com o portal da Imprensa Nacional.
2. Localiza o link da edição completa ("ASSINADO") da data e seção pedidas.
3. Baixa o PDF (com cache local) e extrai o texto.
4. Segmenta o texto para isolar o bloco do MJSP.
5. Extrai os dados conforme a seção (ver mapa abaixo).
6. Gera o documento Word com os resultados.

## 2. Mapa do código

### Configuração (linhas 30–60)
- `AGENT_DIR`, `HOST_BASE`, `RAIZ_REPO` — caminhos portáveis derivados do script.
- `_escolher_diretorio()` — resolve o 1º diretório gravável para `RESULT_DIR` e
  `CACHE_DIR`; no Streamlit Cloud (filesystem somente leitura) cai para o tempdir.
- `PESQUISA_URL`, `START_URL` — endpoints do portal (busca e aquecimento de sessão).
- `SECOES` — cadastro das seções: DOU1 (Seção 1, jornal `515`) e DOU2 (Seção 2, jornal `529`).
- `MINISTERIO_ALVO` — cabeçalho do ministério a ser filtrado.
- `NOISE_RE` — padrões de rodapé/assinatura digital a descartar.
- `MIN_HEADER_RE` / `MJSP_RE` — reconhecem cabeçalhos de ministério e o do MJSP.
- `ACT_HEADER_RE` — reconhece títulos de atos na Seção 1 (alvará, portaria, resolução, etc.).
- `ASSINADO_RE` — extrai o link de download da edição completa da página HTML.
- `UA` — cabeçalho `User-Agent` de navegador (o portal bloqueia bots).

### Rede e download (linhas 113–170)
- `criar_sessao()` — cria `requests.Session` e aquece a sessão acessando `start.action`
  (necessário para receber cookies; o portal retorna 403/500 sem sessão).
- `obter_link_edicao(sessao, secao, data)` — busca a edição da data na seção via `GET`
  em `jornalList.action` com os parâmetros `edicao.*`; devolve o URL do PDF
  (convertendo `&amp;` em `&`), ou `None` se não houver edição.
- `baixar_pdf(url, destino)` — baixa o PDF em streaming; reutiliza o arquivo se o
  cache já existir (> 1 KB).
- `ler_paginas(pdf_path)` / `extrair_texto(pdf_path)` — extraem o texto de todas as
  páginas com `pymupdf`, separando as páginas por `\n\f` (o segundo une as páginas).

### Índices (Sumário) das seções (linhas 171–251)
- `ler_indices(paginas)` — lê, na 1ª página do PDF, as entradas do sumário no formato
  `Nome do índice ... <página inicial>` e devolve a lista ordenada de `{nome, pagina}`.
- `_inicios_paginas(paginas)` / `_linha_cabecalho_indice(...)` — auxiliares que ancoram
  o cabeçalho de cada índice ao texto sem paginação.
- `fatiar_indices(texto, paginas, nomes_selecionados)` — limita o texto aos blocos dos
  índices escolhidos: localiza a linha exata do cabeçalho de cada índice no conteúdo e
  corta do cabeçalho deste até o cabeçalho do índice seguinte (ou fim do documento).

### Segmentação por órgão (linhas 252–275)
- `achar_bloco_mjsp(texto)` — localiza a primeira linha que é exatamente o cabeçalho
  do MJSP e retorna o intervalo (`inicio, fim`) até o cabeçalho do próximo ministério.
- `limpar_linha(linha)` — heurística de ruído: vazio, número isolado (1–3 dígitos) ou
  `NOISE_RE` → `True` (linha a ignorar).

### Seção 1 — atos normativos de segurança pública (linhas 276–332)
- `extrair_atos_secao1(linhas)` — varre o bloco MJSP, agrupa cada ato (título +
  corpo até o próximo título) e mantém apenas os de segurança pública.
- `SEGURANCA_TERMOS` — termos específicos que caracterizam o assunto
  (alvará, munição, arma, tráfico, etc.).
- `eh_ato_seguranca_publica(ato)` — filtro: o título+corpo contém algum termo da lista.
- `resumir_corpo(corpo, max_len=900)` — resume o texto a partir de "resolve:".

### Seção 2 — nomeações/exonerações da Polícia Federal (linhas 333–419)
- `extrair_nomeacoes_secao2(linhas)` — dentro do bloco MJSP, isola o sub-bloco
  `POLÍCIA FEDERAL` (até o próximo órgão), divide em portarias (`DG/PF`, `DGP/PF`,
  `DDG/PF`) e extrai, de cada item de pessoal: número (`Nº`), verbo (Nomear /
  Designar / Exonerar / Dispensar), nome completo (caixa alta) e detalhes.
- `_limpar_resto(texto)` — remove assinaturas e linhas residuais do trecho de detalhes.

### Geração do documento Word (linhas 420–481)
- `gerar_docx(data, resultados)` — monta o relatório com cabeçalho (fonte, data,
  filtro), a Seção 1 em lista numerada com bullets, e a Seção 2 em tabela
  5 colunas (Nº, Ato, Ação, Nome, Detalhes).

### Fluxo principal (linhas 482–551)
- `processar_secao(sessao, secao, data, indices=None)` — orquestra uma seção: link →
  download → páginas → (opcional) fatia pelos índices selecionados → bloco MJSP →
  extração (devolve `{"secao1": [...]}` ou `{"secao2": [...]}`).
- `main()` — parseia argumentos de linha de comando, valida a data, processa as
  seções pedidas (com filtro opcional de índices) e gera o relatório.

## 3. Uso

```text
python D:\opencode_desktop\Agents\DOU_Research_Agent\agent.py [--data DD/MM/AAAA] [--secoes 1,2] [--indices "Nome do índice 1;Nome do índice 2"]
```

| Argumento | Padrão | Descrição |
| --- | --- | --- |
| `--data` | hoje | Data da edição no formato `DD/MM/AAAA`. |
| `--secoes` | `1,2` | Seções a processar, separadas por vírgula. |
| `--indices` | (vazio) | Índices (entradas do sumário do DOU) a pesquisar, separados por `;`. A busca fica limitada a esses blocos; aplica-se a todas as seções processadas. |

Exemplos:

```text
python ...\agent.py                          # hoje, seções 1 e 2
python ...\agent.py --data 08/09/2026        # data específica, seções 1 e 2
python ...\agent.py --data 08/09/2026 --secoes 2   # só nomeações/exonerações PF
python ...\agent.py --data 08/09/2026 --indices "Ministério da Justiça e Segurança Pública"
                                            # só o bloco do MJSP (Seções 1 e 2)
```

## 4. Exemplo de saída (edição 08/09/2026)

```text
--- Seção 1 (Diário Oficial da União - Seção 1) ---
  [cache] 2026_09_08_ASSINADO_do1.pdf já existe (32310 KB)
  [ok] Bloco MJSP: linhas 21211..26155 (4945 linhas)
  [ok] 21 atos normativos localizados

--- Seção 2 (Diário Oficial da União - Seção 2) ---
  [cache] 2026_09_08_ASSINADO_do2.pdf já existe (21679 KB)
  [ok] Bloco MJSP: linhas 12274..12491 (218 linhas)
  [ok] 7 nomeações/exonerações PF localizadas

Gerando relatório Word...
Relatório gerado: D:\opencode_desktop\resultado\Relatorio_DOU_MJSP_2026_09_08.docx
```

## 5. Aplicação Streamlit

Há uma interface gráfica em `app.py` que reutiliza as funções de `agent.py`.
Ela permite escolher a data, **carregar os índices (sumário) da edição** e
selecionar quais índices pesquisar, pela lateral; executar a pesquisa com log de
progresso; visualizar os resultados em abas (Seção 1 em expanders e Seção 2 em
tabela) e baixar o relatório Word.

Executar (a partir de qualquer pasta):

```text
streamlit run D:\opencode_desktop\Agents\DOU_Research_Agent\app.py
```

Fluxo na barra lateral:
1. Escolha a data da edição.
2. Clique em **1️⃣ Carregar índices da edição** — baixa os PDFs das Seções 1 e 2
   (em cache) e lê o sumário de cada uma.
3. No seletor **Índices da edição a pesquisar**, marque os índices desejados
   (ex.: `Seção 2 · Ministério da Justiça e Segurança Pública`) ou `(edição
   inteira)` para varrer a seção completa. Todos os índices das Seções 1 e 2 são
   opções selecionáveis.
4. Clique em **2️⃣ Executar pesquisa** — a busca fica limitada aos índices
   selecionados.

Opções adicionais: "Forçar novo download (ignorar cache)", que apaga o PDF da
data antes de baixar novamente.

## 6. Publicação no GitHub e deploy no Streamlit Community Cloud

O código vive no repositório público `https://github.com/agentdoust/dou-research-agent`
(branch `main`). A raiz do repositório contém `requirements.txt` (dependências do
deploy) e o app fica em `Agents/DOU_Research_Agent/app.py`.

### Deploy no Streamlit Community Cloud

1. Acesse `https://share.streamlit.io/` e faça login com a conta GitHub
   `agentdoust` (a autenticação é via OAuth do GitHub — mesma conta).
2. **Novo app → Crie um app**: repositório `agentdoust/dou-research-agent`,
   branch `main`, arquivo principal `Agents/DOU_Research_Agent/app.py`.
3. Clique em **Deploy**. O cloud instala o `requirements.txt` da raiz e executa o
   app.
4. Resultado: `https://dou-research-agent.streamlit.app`.

### Comportamento no cloud

- O filesystem do contêiner é **somente leitura**; os diretórios de cache e de
  resultados caem automaticamente para o tempdir (`_escolher_diretorio` em
  `agent.py:42`), então cada sessão baixa os PDFs sob demanda.
- Segredos (se um dia forem necessários) vão em **Settings → Secrets** do app
  (guardiões em `.streamlit/secrets.toml`, ignorado pelo Git).

## 7. Estrutura de arquivos

```text
D:\opencode_desktop\          (raiz do repositório GitHub)
├─ .gitignore                 # exclui downloads, resultado, arquivos_txt, etc.
├─ requirements.txt           # dependências do deploy (Streamlit Cloud)
├─ .streamlit\config.toml     # tema/apresentação do app (opcional)
├─ arquivos_txt\              # notas e credenciais — NÃO versionado
│  └─ CRDS                    # credenciais GitHub/Streamlit (gitignored)
├─ Agents\DOU_Research_Agent\
│  ├─ agent.py                # lógica do agente (linha de comando)
│  ├─ app.py                  # interface Streamlit
│  ├─ README.md
│  └─ downloads\              # cache dos PDFs (<YYYY_MM_DD>\<arquivo>.pdf) — ignorado
├─ Fonte\                     # PDFs de edições fornecidas manualmente (consulta) — ignorado
└─ resultado\                 # relatórios Word gerados — ignorado
```

## 8. Observações técnicas

- **Sessão obrigatória**: o portal exige cookies; por isso `criar_sessao()` chama
  `start.action` antes da busca. `POST` em `jornalList.action` é bloqueado (403);
  usa-se `GET` com os parâmetros `edicao.*`.
- **Cache**: o download é reaproveitado se o arquivo já existir (> 1 KB); use a
  opção "Forçar novo download" na app (ou remova `downloads\<data>\`) para refazer.
- **Acentos/console**: no Windows/PowerShell o console pode exibir acentos
  corrompidos (cp1252), mas os arquivos e o Word preservam UTF-8.
- **Dependências**: `pip install requests pymupdf python-docx streamlit` (Python 3.13+).
  Para o deploy, `requirements.txt` na raiz do repositório.
- **Diretórios no cloud**: `resultado/` e `downloads/` só são graváveis em máquina
  local; no Streamlit Cloud o agente os remapeia para o tempdir automaticamente.

## 9. Limitações conhecidas

- A segmentação assume o cabeçalho do MJSP exatamente na forma
  `Ministério da Justiça e Segurança Pública`; variações do nome podem não casar.
- O filtro da Seção 1 usa uma lista fixa de termos; atos que não citem esses termos
  podem ficar de fora (e as heurísticas de "tráfico de" exigem a preposição).
- A extração da Seção 2 foca nas portarias `DG/PF`, `DGP/PF` e `DDG/PF`.