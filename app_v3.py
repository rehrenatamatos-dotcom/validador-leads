# Validador de Leads v3 — Soluções Industriais
# v2 (busca automática no Metabase) + dashboard HTML de resultados para download.

import csv
import io
import json
import os
import re
import time
from datetime import date, datetime, timedelta

import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# Cores de destaque no Excel por status
FILL_DENTRO = PatternFill("solid", fgColor="C6EFCE")   # verde
FILL_FORA = PatternFill("solid", fgColor="FFC7CE")     # vermelho
FILL_ABERTO = PatternFill("solid", fgColor="FFEB9C")   # âmbar
FILL_CABECALHO = PatternFill("solid", fgColor="0C447C")  # azul escuro

# Provedores gratuitos (APIs compatíveis com OpenAI). O app escolhe automaticamente:
# se houver CEREBRAS_API_KEY nos Secrets usa Cerebras (cota diária bem maior);
# senão usa Groq. Um serve de reserva do outro quando ambas as chaves existem.
# Cada provedor tem uma LISTA de modelos candidatos. Se um estiver descontinuado
# (404), o app tenta o próximo automaticamente — nunca quebra por um nome só.
# Modelos confirmados ATIVOS (consultados na doc oficial de cada provedor, jul/2026).
# Cada provedor tem uma lista de candidatos: se um cair (404), tenta o próximo.
PROVEDORES = {
    "cerebras": {
        "url": "https://api.cerebras.ai/v1/chat/completions",
        "chave": "CEREBRAS_API_KEY",
        "modelos": ["llama-3.3-70b", "llama3.1-8b", "gpt-oss-120b"],
    },
    "groq": {
        "url": "https://api.groq.com/openai/v1/chat/completions",
        "chave": "GROQ_API_KEY",
        "modelos": ["openai/gpt-oss-20b", "llama-3.1-8b-instant",
                    "llama-3.3-70b-versatile", "openai/gpt-oss-120b"],
    },
}

# Opções do seletor de IA na tela. Cada opção define (provedor, modelo preferido).
# Todos os modelos abaixo foram confirmados ativos na documentação oficial.
# Cerebras tem cota diária muito maior (1M tokens/dia) — melhor para uso intenso.
MODELOS_ESCOLHA = {
    "Automático (recomendado)": None,
    "Rápido · Cerebras Llama 3.1 8B": ("cerebras", "llama3.1-8b"),
    "Rápido · Groq GPT-OSS 20B": ("groq", "openai/gpt-oss-20b"),
    "Equilíbrio · Cerebras Llama 3.3 70B": ("cerebras", "llama-3.3-70b"),
    "Equilíbrio · Groq Llama 3.3 70B": ("groq", "llama-3.3-70b-versatile"),
    "Preciso · Groq GPT-OSS 120B": ("groq", "openai/gpt-oss-120b"),
}

TAMANHO_LOTE = 20
MAX_TENTATIVAS = 5
LIMITE_FONTE = 4000
LIMITE_PERFIL = 9000
STATUS_VALIDOS = {"Dentro do foco", "Fora do foco", "Aberto"}

CARD_ORCAMENTOS = 47   # question do Metabase: base de orçamentos por chave única
CARD_BRIEFING = 286    # question do Metabase: briefing do cliente por chave única
CARD_ANUNCIOS = 185    # question do Metabase: anúncios por chave única

AZUL = "#185FA5"
AZUL_ESCURO = "#0C447C"
AZUL_CLARO = "#E6F1FB"

ARQUIVO_HISTORICO = "historico.json"
PASTA_RESULTADOS = "resultados"

PROMPT_SISTEMA = """Você é um analista de qualidade de leads de uma plataforma de geração de leads B2B.
Sua tarefa: para cada lead abaixo, decidir se ele está DENTRO ou FORA do foco do cliente descrito no perfil, ou se está ABERTO (mensagem sem informação suficiente para avaliar).

ANTES DE CLASSIFICAR: identifique no perfil O QUE o cliente vende — um PRODUTO (ex. máquinas, equipamentos) ou um SERVIÇO (ex. corte sob medida, usinagem para terceiros). Essa distinção é o critério mais importante.

Critérios (avalie em conjunto, nenhum sozinho decide):
1. Produto vs. serviço. Se o cliente VENDE MÁQUINAS e o lead quer CONTRATAR o serviço (ex. "preciso cortar 50 chapas", "orçamento para corte de peças"), é "Fora do foco" — mesmo que a mensagem seja tecnicamente detalhada (material, medidas, CNPJ). Especificidade técnica NÃO transforma um pedido de serviço em lead de máquina. O inverso também vale (cliente presta serviço e lead quer comprar máquina). Salvo se o perfil disser que o cliente atende ambos.
2. Modalidade compatível. Pedidos de assistência técnica/manutenção, aluguel de máquina, ou peças/componentes avulsos são "Fora do foco" quando o cliente vende equipamentos novos — salvo indicação contrária no perfil ou nas regras específicas.
3. Material/produto compatível com o portfólio do cliente. Se o cliente trabalha metal e o lead pede madeira/tecido/PVC, é forte sinal de "Fora do foco", mesmo que o serviço seja o mesmo.
4. B2B vs. uso pessoal. Pedidos claramente domésticos/pontuais de pessoa física pesam para "Fora do foco" quando o cliente atende indústria/B2B.
5. Especificidade técnica. Medidas, normas, quantidade definida, nome de empresa/CNPJ pesam para "Dentro do foco" — mas SOMENTE quando o pedido é da modalidade certa (ver critério 1).
6. Sinais de ruído. Teste interno (QA, e-mails de qualidade), spam, concorrente se oferecendo, marca/modelo que o cliente não vende, ou lead avisando que já comprou em outro lugar = "Fora do foco" independente do produto.
7. Mensagem inteiramente em inglês. Se a mensagem do lead estiver totalmente escrita em inglês (ex. "Dear Sir/Madam, we are interested in your products..."), classifique como "Fora do foco" — são tipicamente bots ou contatos genéricos internacionais fora do público-alvo. Isso vale mesmo que a mensagem pareça pedir um produto do cliente. NÃO se aplica a mensagens em português que contenham apenas termos técnicos ou nomes de produto em inglês (ex. "máquina laser CO2", "new laser nli390") — essas continuam sendo avaliadas normalmente.
8. Regras específicas do cliente (se fornecidas no perfil) têm prioridade sobre os critérios gerais.

Regras de saída:
- STATUS deve ser EXATAMENTE um destes: "Dentro do foco", "Fora do foco", "Aberto".
- MOTIVO: uma frase objetiva em português citando a evidência da própria mensagem. O motivo deve justificar o STATUS escolhido, não outro.
- Mensagens vagas demais para julgar (ex. apenas "aço inox", apenas "me manda o e-mail", uma palavra solta sem contexto de compra) = "Aberto". NUNCA marque "Dentro do foco" sem evidência de interesse na modalidade certa (compra do produto que o cliente vende).
- Peças, componentes e insumos avulsos (ex. fonte, tubo de laser, lentes) = "Fora do foco" quando o cliente vende máquinas completas, salvo indicação contrária no perfil.
- Mensagens idênticas ou quase idênticas (mesmo texto em vários leads) DEVEM receber exatamente a mesma classificação e o mesmo motivo — revise antes de responder.
- Mensagem inteiramente em inglês = "Fora do foco" (ver critério 7), com motivo indicando que é mensagem em inglês / provável bot.
- Responda SOMENTE com um objeto JSON: {"resultados": [{"id": "...", "status": "...", "motivo": "..."}]} — um item por lead, na mesma ordem."""


# ---------- Histórico ----------

def carregar_historico():
    try:
        with open(ARQUIVO_HISTORICO, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def salvar_no_historico(registro, xlsx_bytes=None, dash_bytes=None):
    historico = carregar_historico()
    historico.insert(0, registro)
    historico = historico[:200]
    try:
        with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False)
        os.makedirs(PASTA_RESULTADOS, exist_ok=True)
        rid = registro.get("id", "")
        if rid and xlsx_bytes:
            with open(os.path.join(PASTA_RESULTADOS, f"{rid}.xlsx"), "wb") as f:
                f.write(xlsx_bytes)
        if rid and dash_bytes:
            with open(os.path.join(PASTA_RESULTADOS, f"{rid}.html"), "wb") as f:
                f.write(dash_bytes)
    except Exception:
        pass


def excluir_do_historico(rid):
    historico = [h for h in carregar_historico() if h.get("id") != rid]
    try:
        with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False)
        for ext in (".csv", ".xlsx", ".html"):
            caminho = os.path.join(PASTA_RESULTADOS, f"{rid}{ext}")
            if os.path.exists(caminho):
                os.remove(caminho)
    except Exception:
        pass


def ler_resultado_salvo(rid, ext):
    try:
        with open(os.path.join(PASTA_RESULTADOS, f"{rid}{ext}"), "rb") as f:
            return f.read()
    except Exception:
        return None


# ---------- Metabase ----------

def secret(nome, padrao=""):
    try:
        return str(st.secrets.get(nome, padrao)).strip()
    except Exception:
        return os.environ.get(nome, padrao).strip()


def cabecalhos_metabase():
    """Autentica no Metabase por API key ou por login/senha (sessão)."""
    api_key = secret("METABASE_API_KEY")
    if api_key:
        return {"X-API-KEY": api_key}
    if "mb_sessao" in st.session_state:
        return {"X-Metabase-Session": st.session_state["mb_sessao"]}
    usuario, senha = secret("METABASE_USER"), secret("METABASE_PASSWORD")
    if not (usuario and senha):
        return None
    r = requests.post(
        secret("METABASE_URL").rstrip("/") + "/api/session",
        json={"username": usuario, "password": senha},
        timeout=30,
    )
    r.raise_for_status()
    st.session_state["mb_sessao"] = r.json()["id"]
    return {"X-Metabase-Session": st.session_state["mb_sessao"]}


def consultar_question(card_id, parametros):
    """Baixa o resultado CSV de uma question do Metabase com parâmetros."""
    url = secret("METABASE_URL").rstrip("/") + f"/api/card/{card_id}/query/csv"
    headers = cabecalhos_metabase()
    if headers is None:
        raise RuntimeError(
            "Credenciais do Metabase não configuradas. Adicione nos Secrets: "
            "METABASE_URL e (METABASE_API_KEY ou METABASE_USER + METABASE_PASSWORD)."
        )
    params_mb = [
        {"type": tipo, "target": ["variable", ["template-tag", nome]], "value": valor}
        for nome, valor, tipo in parametros
    ]
    r = requests.post(url, headers=headers, data={"parameters": json.dumps(params_mb)}, timeout=120)
    if r.status_code == 401 and "mb_sessao" in st.session_state:
        del st.session_state["mb_sessao"]           # sessão expirada → renova e tenta de novo
        headers = cabecalhos_metabase()
        r = requests.post(url, headers=headers, data={"parameters": json.dumps(params_mb)}, timeout=120)
    r.raise_for_status()
    return r.content.decode("utf-8-sig", errors="replace")


# ---------- Fontes auxiliares ----------

def csv_anuncios_para_texto(conteudo_csv, limite=LIMITE_FONTE):
    """Transforma o CSV de anúncios em uma lista compacta de nomes."""
    linhas = list(csv.reader(io.StringIO(conteudo_csv)))
    if len(linhas) < 2:
        return ""
    itens = []
    for r in linhas[1:]:
        valores = [c.strip() for c in r if c.strip()]
        if valores:
            itens.append(" | ".join(valores))
    return "\n".join(itens)[:limite]


def extrair_nome_empresa(conteudo_csv):
    """Procura o nome do cliente em qualquer CSV do Metabase que tenha essa coluna."""
    try:
        linhas = list(csv.reader(io.StringIO(conteudo_csv)))
        if len(linhas) < 2:
            return None
        h = [c.strip().lower() for c in linhas[0]]
        for cand in ("nome fantasia", "nome da empresa", "empresa"):
            if cand in h:
                v = linhas[1][h.index(cand)].strip()
                if v:
                    return v
    except Exception:
        pass
    return None


def buscar_site(url, limite=LIMITE_FONTE):
    try:
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        html = r.text
    except Exception:
        return ""
    html = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", html)
    texto = re.sub(r"(?s)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", texto).strip()[:limite]


def csv_briefing_para_texto(conteudo_csv, limite=LIMITE_FONTE):
    """Converte o CSV do briefing (colunas longas) em texto legível para a IA."""
    linhas = list(csv.reader(io.StringIO(conteudo_csv)))
    if len(linhas) < 2:
        return ""
    h = linhas[0]
    blocos = []
    for r in linhas[1:]:
        campos = [f"{h[j]}: {r[j].strip()}" for j in range(min(len(h), len(r))) if r[j].strip()]
        blocos.append("\n".join(campos))
    return ("\n\n--- produto/linha seguinte ---\n\n".join(blocos))[:limite]


# ---------- Dashboard HTML ----------

MODELO_DASH = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relatório de Validação — __EMPRESA__</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', Arial, sans-serif; color: #EAF3FC; padding: 24px; min-height: 100vh;
    background:
      linear-gradient(rgba(12, 68, 124, 0.86), rgba(10, 58, 107, 0.90)),
      url('https://images.unsplash.com/photo-1513828583688-c52646db42da?w=1600&q=60') center / cover fixed no-repeat;
  }
  .wrap { max-width: 900px; margin: 0 auto; }
  .vidro {
    background: rgba(255, 255, 255, 0.13);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.32);
    border-radius: 16px;
  }
  .topo { padding: 22px 28px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
  .topo h1 { color: #fff; font-size: 20px; font-weight: 600; }
  .topo p { color: #B5D4F4; font-size: 13px; margin-top: 4px; }
  .topo .marca { color: #B5D4F4; font-size: 12px; text-align: right; }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin: 14px 0; }
  .card { padding: 16px; text-align: center; }
  .card .rotulo { font-size: 12px; color: #B5D4F4; margin-bottom: 4px; }
  .card .valor { font-size: 30px; font-weight: 600; }
  .card .sub { font-size: 12px; color: rgba(234,243,252,0.7); margin-top: 2px; }
  .azul { color: #ffffff; } .verde { color: #9FE1A5; } .vermelho { color: #F5A9A9; } .ambar { color: #FAD98F; }
  .paineis { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .painel { padding: 18px; }
  .painel h2 { font-size: 14px; color: #ffffff; margin-bottom: 12px; }
  .painel h2 .pin { font-size: 11px; padding: 2px 8px; border-radius: 20px; vertical-align: middle; margin-left: 6px; }
  .pin-verde { background: rgba(159,225,165,0.25); color: #C6EFCE; }
  .pin-verm { background: rgba(245,169,169,0.25); color: #F5A9A9; }
  .tab-leads { width: 100%; border-collapse: collapse; font-size: 12px; }
  .tab-leads th { text-align: left; color: #B5D4F4; font-weight: 600; padding: 4px 6px; border-bottom: 1px solid rgba(255,255,255,0.2); }
  .tab-leads td { padding: 5px 6px; border-bottom: 1px solid rgba(255,255,255,0.1); color: #EAF3FC; }
  .anuncio-linha { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; font-size: 12px; }
  .anuncio-linha .nome { width: 150px; color: #EAF3FC; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .barra-fundo { flex: 1; height: 8px; background: rgba(255,255,255,0.12); border-radius: 5px; overflow: hidden; }
  .barra-cheia { height: 100%; background: #F5A9A9; border-radius: 5px; }
  .anuncio-linha .qtd { width: 22px; text-align: right; color: #B5D4F4; }
  .rodape { text-align: center; color: rgba(234,243,252,0.6); font-size: 11px; margin-top: 18px; }
  @media (max-width: 700px) { .paineis { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="wrap">
  <div class="topo vidro">
    <div>
      <h1>Relatório de Validação de Leads</h1>
      <p>__EMPRESA__ &nbsp;·&nbsp; chave __CHAVE__ &nbsp;·&nbsp; __PERIODO__</p>
    </div>
    <div class="marca">Soluções Industriais<br>gerado em __GERADO__</div>
  </div>

  <div class="cards">
    <div class="card vidro"><div class="rotulo">Total de leads</div><div class="valor azul">__TOTAL__</div><div class="sub">no período</div></div>
    <div class="card vidro"><div class="rotulo">Dentro do foco</div><div class="valor verde">__PCT_DENTRO__%</div><div class="sub">__N_DENTRO__ leads</div></div>
    <div class="card vidro"><div class="rotulo">Fora do foco</div><div class="valor vermelho">__PCT_FORA__%</div><div class="sub">__N_FORA__ leads</div></div>
    <div class="card vidro"><div class="rotulo">Aberto</div><div class="valor ambar">__PCT_ABERTO__%</div><div class="sub">__N_ABERTO__ leads</div></div>
  </div>

  <div class="paineis">
    <div class="painel vidro">
      <h2>Distribuição por status</h2>
      <canvas id="rosca" height="220"></canvas>
    </div>
    <div class="painel vidro">
      <h2>Anúncios que mais geraram leads fora do foco</h2>
      __LINHAS_ANUNCIOS__
    </div>
  </div>

  <div class="paineis" style="margin-top:12px;">
    <div class="painel vidro">
      <h2>Melhores leads <span class="pin pin-verde">dentro do foco</span></h2>
      <table class="tab-leads">
        <tr><th>ID</th><th>Nome</th><th>E-mail</th></tr>
        __LINHAS_MELHORES__
      </table>
    </div>
    <div class="painel vidro">
      <h2>Piores leads <span class="pin pin-verm">fora do foco</span></h2>
      <table class="tab-leads">
        <tr><th>ID</th><th>Nome</th><th>E-mail</th></tr>
        __LINHAS_PIORES__
      </table>
    </div>
  </div>

  <p class="rodape">Validador de Leads · Soluções Industriais · uso interno</p>
</div>
<script>
const CORES = ["#639922", "#E24B4A", "#EF9F27"];
const ROTULOS = ["Dentro do foco", "Fora do foco", "Aberto"];
const DADOS = [__N_DENTRO__, __N_FORA__, __N_ABERTO__];
const CLARO = "#EAF3FC";
new Chart(document.getElementById("rosca"), {
  type: "doughnut",
  data: { labels: ROTULOS, datasets: [{ data: DADOS, backgroundColor: CORES, borderWidth: 2, borderColor: "rgba(255,255,255,0.4)" }] },
  options: { plugins: { legend: { position: "bottom", labels: { color: CLARO } } }, cutout: "58%" }
});
</script>
</body>
</html>"""


def gerar_xlsx(cabecalho, registros, leads, classificacoes):
    """Gera um Excel com as colunas originais + STATUS + MOTIVO, pintando cada
    linha por status: verde (dentro), vermelho (fora), âmbar (aberto)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads validados"
    colunas = cabecalho + ["STATUS", "MOTIVO"]
    ws.append(colunas)
    for cel in ws[1]:
        cel.fill = FILL_CABECALHO
        cel.font = Font(color="FFFFFF", bold=True)
    for i, r in enumerate(registros):
        c = classificacoes.get(leads[i]["id"], {
            "status": "Aberto", "motivo": "Não classificado pela IA — revisar manualmente.",
        })
        linha = list(r) + [c["status"], c["motivo"]]
        ws.append(linha)
        fill = {"Dentro do foco": FILL_DENTRO, "Fora do foco": FILL_FORA}.get(c["status"], FILL_ABERTO)
        for cel in ws[ws.max_row]:
            cel.fill = fill
    # larguras aproximadas para leitura
    for col in ws.columns:
        larg = min(60, max(12, max((len(str(c.value)) if c.value else 0) for c in col) + 2))
        ws.column_dimensions[col[0].column_letter].width = larg
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def gerar_dashboard_html(empresa, chave, periodo, total, contagem,
                         melhores=None, piores=None, anuncios_ruins=None):
    def pct(n):
        return str(round(100 * n / total)) if total else "0"

    def linhas_leads(lista):
        if not lista:
            return "<tr><td colspan='3' style='color:rgba(234,243,252,0.6);'>Nenhum lead nesta categoria.</td></tr>"
        out = ""
        for ld in lista:
            out += (f"<tr><td>{ld['id']}</td><td>{ld['nome']}</td>"
                    f"<td style='font-size:11px;'>{ld['email']}</td></tr>")
        return out

    def linhas_anuncios(lista):
        if not lista:
            return "<p style='color:rgba(234,243,252,0.6); font-size:12px;'>Sem dados de anúncio.</p>"
        maior = max(qtd for _, qtd in lista) or 1
        out = ""
        for nome, qtd in lista:
            largura = round(100 * qtd / maior)
            out += (f"<div class='anuncio-linha'><div class='nome'>{nome}</div>"
                    f"<div class='barra-fundo'><div class='barra-cheia' style='width:{largura}%'></div></div>"
                    f"<div class='qtd'>{qtd}</div></div>")
        return out

    html = MODELO_DASH
    trocas = {
        "__EMPRESA__": empresa,
        "__CHAVE__": chave,
        "__PERIODO__": periodo,
        "__GERADO__": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "__TOTAL__": str(total),
        "__PCT_DENTRO__": pct(contagem["Dentro do foco"]),
        "__PCT_FORA__": pct(contagem["Fora do foco"]),
        "__PCT_ABERTO__": pct(contagem["Aberto"]),
        "__N_DENTRO__": str(contagem["Dentro do foco"]),
        "__N_FORA__": str(contagem["Fora do foco"]),
        "__N_ABERTO__": str(contagem["Aberto"]),
        "__LINHAS_MELHORES__": linhas_leads(melhores),
        "__LINHAS_PIORES__": linhas_leads(piores),
        "__LINHAS_ANUNCIOS__": linhas_anuncios(anuncios_ruins),
    }
    for k, v in trocas.items():
        html = html.replace(k, v)
    return html


# ---------- IA (Cerebras ou Groq, compatíveis com OpenAI) ----------

def provedores_ativos():
    """Retorna a lista de provedores com chave configurada (Cerebras primeiro)."""
    return [n for n in ("cerebras", "groq") if secret(PROVEDORES[n]["chave"])]


def _extrair_lista(parsed):
    if isinstance(parsed, list):
        return parsed
    for v in parsed.values():
        if isinstance(v, list):
            return v
    raise ValueError("Resposta sem lista de resultados.")


def _tentar_modelo(url, api_key, modelo, conteudo_user):
    """Uma chamada a um modelo específico. Retorna (lista, None) em sucesso,
    ('404', None) se o modelo não existe, ('cota', horas) se a cota diária estourou,
    (None, excecao) em outros erros."""
    corpo = {
        "model": modelo,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": conteudo_user},
        ],
    }
    ultima = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = requests.post(url, json=corpo, headers={"Authorization": f"Bearer {api_key}"}, timeout=120)
            if r.status_code in (400, 404):
                return "404", None                      # modelo indisponível: tenta o próximo
            if r.status_code == 429 or r.status_code >= 500:
                try:
                    espera = float(r.headers.get("retry-after", 0))
                except (TypeError, ValueError):
                    espera = 0
                if espera > 300:
                    return "cota", espera
                time.sleep(min(60, espera + 1) if espera else min(30, 5 * tentativa))
                continue
            r.raise_for_status()
            texto = r.json()["choices"][0]["message"]["content"]
            return _extrair_lista(json.loads(texto)), None
        except Exception as e:
            ultima = e
            time.sleep(2)
    return None, ultima


def chamar_ia(perfil, lote, ordem, modelo_forcado=None):
    """Tenta os provedores em ordem; dentro de cada um, tenta a lista de modelos candidatos
    (começando pelo escolhido). Pula modelos indisponíveis (404) e troca de provedor na cota."""
    leads_texto = "\n\n".join(
        f"LEAD id={l['id']}\nMensagem: {l['mensagem']}\nContexto extra: {l['extra']}"
        for l in lote
    )
    conteudo_user = f"PERFIL DO CLIENTE:\n{perfil}\n\nLEADS A CLASSIFICAR:\n{leads_texto}"
    ultima = None
    esgotados = []
    for i, nome in enumerate(ordem):
        cfg = PROVEDORES[nome]
        api_key = secret(cfg["chave"])
        # ordem de modelos: o escolhido (se for deste provedor) primeiro, depois os candidatos
        modelos = list(cfg["modelos"])
        if modelo_forcado and i == 0 and modelo_forcado in modelos:
            modelos.remove(modelo_forcado)
            modelos.insert(0, modelo_forcado)
        elif modelo_forcado and i == 0:
            modelos.insert(0, modelo_forcado)

        cota_estourou = False
        for modelo in modelos:
            resultado, err = _tentar_modelo(cfg["url"], api_key, modelo, conteudo_user)
            if isinstance(resultado, list):
                return resultado
            if resultado == "404":
                continue                                # modelo indisponível: próximo candidato
            if resultado == "cota":
                esgotados.append(nome)
                cota_estourou = True
                break
            ultima = err                                # erro genérico: tenta próximo modelo
        if cota_estourou:
            continue                                    # próximo provedor
    if esgotados and len(esgotados) == len(ordem):
        raise RuntimeError(
            "Cota diária esgotada em todos os provedores de IA configurados "
            f"({', '.join(esgotados)}). Renova às 21h (horário de Brasília), ou adicione "
            "outra chave (CEREBRAS_API_KEY ou GROQ_API_KEY) nos Secrets para continuar agora."
        )
    if ultima:
        raise ultima
    raise RuntimeError(
        "Nenhum modelo respondeu nos provedores configurados. "
        "Confira se as chaves de IA nos Secrets estão corretas e ativas."
    )


# ---------- Interface ----------

st.set_page_config(page_title="Validador de Leads v3", page_icon="✅", layout="wide")

FOTO_INDUSTRIA = "https://images.unsplash.com/photo-1513828583688-c52646db42da?w=1900&q=70"


def obter_nome_analista():
    """Nome do analista: Secrets (ANALISTA_NOME, fixo e definitivo) > link salvo
    (?analista=... na URL) > sessão atual. Cada analista tem seu próprio app, então
    o mais simples e permanente é configurar ANALISTA_NOME nos Secrets dele."""
    nome = secret("ANALISTA_NOME")
    if nome:
        return nome
    try:
        da_url = st.query_params.get("analista", "")
    except Exception:
        da_url = ""
    return da_url.strip() or st.session_state.get("nome_analista", "")


nome_analista = obter_nome_analista()

if not nome_analista:
    st.markdown(f"""
    <style>
      html, body, [data-testid="stAppViewContainer"] {{ background: #0B1622; }}
      [data-testid="stHeader"] {{ background: transparent; }}
      .block-container {{ max-width: 480px; padding-top: 8vh; }}
      .boasvindas {{
        background: linear-gradient(160deg, rgba(6,20,36,0.80), rgba(12,68,124,0.72)),
          url('{FOTO_INDUSTRIA}') center/cover no-repeat;
        border-radius: 22px; padding: 44px 40px 34px; text-align: center;
      }}
      .selo-glass {{ display: inline-block; background: rgba(255,255,255,0.14); border: 1px solid rgba(255,255,255,0.3);
        color: #EAF3FC; font-size: 11.5px; padding: 4px 14px; border-radius: 20px; margin-bottom: 16px; }}
      .boasvindas h1 {{ color: #fff; font-size: 21px; font-weight: 700; margin-bottom: 8px; }}
      .boasvindas p {{ color: #D9EAFB; font-size: 13px; margin-bottom: 4px; line-height: 1.55; }}
      div[data-baseweb="input"] {{
        background: rgba(255,255,255,0.14) !important; border: 1px solid rgba(255,255,255,0.38) !important;
        border-radius: 12px !important;
      }}
      div[data-baseweb="input"] input {{
        background: transparent !important; color: #ffffff !important; text-align: center;
      }}
      div[data-baseweb="input"] input::placeholder {{ color: rgba(255,255,255,0.55) !important; }}
      .stButton > button {{
        background: rgba(255,255,255,0.94) !important; color: {AZUL_ESCURO} !important; border: none !important;
        border-radius: 12px !important; font-weight: 700 !important;
      }}
    </style>
    <div class="boasvindas">
      <div class="selo-glass">Soluções Industriais</div>
      <h1>Bem-vindo(a) ao Validador de Leads</h1>
      <p>Antes de começar, como você se chama?</p>
    </div>
    """, unsafe_allow_html=True)
    nome_digitado = st.text_input(
        "Seu nome", key="f_nome_boas", label_visibility="collapsed", placeholder="Digite seu nome",
    )
    if st.button("Continuar", type="primary", use_container_width=True):
        if nome_digitado.strip():
            st.session_state["nome_analista"] = nome_digitado.strip()
            st.query_params["analista"] = nome_digitado.strip()
            st.rerun()
        else:
            st.warning("Digite seu nome para continuar.")
    st.caption(
        "Dica: depois de continuar, salve o link da barra de endereço nos favoritos — "
        "assim o app já abre com o seu nome da próxima vez. Para algo definitivo, "
        "peça para adicionar ANALISTA_NOME nos Secrets do seu app."
    )
    st.stop()

st.session_state.setdefault("nome_analista", nome_analista)

st.markdown(f"""
<style>
  html, body, [data-testid="stAppViewContainer"] {{ background: #0B1622; }}
  [data-testid="stHeader"] {{ background: transparent; }}

  .block-container {{ max-width: 1180px; padding-top: 1.4rem; padding-bottom: 3rem; }}

  .block-container label, .block-container [data-testid="stWidgetLabel"] p {{
    color: #D9EAFB !important; font-weight: 600 !important; font-size: 0.82rem !important;
  }}
  .block-container p, .block-container .stMarkdown {{ color: #AFCBE8; }}

  .topbar-vidro {{ display: flex; align-items: center; justify-content: space-between; padding: 2px 4px 18px; }}
  .topbar-vidro .marca {{ display: flex; align-items: center; gap: 10px; color: #fff; font-weight: 700; font-size: 14px; }}
  .topbar-vidro .marca .quad {{
    width: 26px; height: 26px; border-radius: 7px; background: {AZUL};
    display: flex; align-items: center; justify-content: center; font-size: 11px;
  }}
  .saudacao {{ color: #EAF3FC; font-size: 13.5px; }}
  .saudacao b {{ color: #fff; }}

  .hero {{
    background: linear-gradient(160deg, rgba(6,20,36,0.72), rgba(12,68,124,0.58)),
      url('{FOTO_INDUSTRIA}') center/cover no-repeat;
    border-radius: 20px; padding: 34px 40px; margin-bottom: 20px;
  }}
  .hero .badge {{
    display: inline-block; font-size: 0.72rem; font-weight: 600; letter-spacing: 0.02em;
    color: #fff; background: rgba(255,255,255,0.16); border: 1px solid rgba(255,255,255,0.32);
    padding: 3px 12px; border-radius: 999px; margin-bottom: 12px;
  }}
  .hero h1 {{ color: #ffffff; font-size: clamp(1.5rem, 3vw, 2rem); font-weight: 700; margin: 0 0 8px; }}
  .hero p {{ color: #D6E8FB; font-size: 0.95rem; line-height: 1.5; max-width: 540px; margin: 0; }}

  /* Painéis com borda (formulário e histórico) ganham a estética de vidro */
  div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.16) !important;
    border-radius: 18px !important;
    backdrop-filter: blur(16px);
  }}

  /* Campos em vidro — no baseweb (biblioteca por baixo do Streamlit), o "quadrado"
     visível com fundo e borda é o div[data-baseweb=...], não o input/textarea em si. */
  div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div {{
    background: rgba(255,255,255,0.10) !important; border: 1px solid rgba(255,255,255,0.26) !important;
    border-radius: 12px !important;
  }}
  div[data-baseweb="input"] input, div[data-baseweb="textarea"] textarea,
  div[data-baseweb="select"] div, div[data-baseweb="select"] span {{
    background: transparent !important; color: #ffffff !important;
  }}
  div[data-baseweb="input"] input::placeholder, div[data-baseweb="textarea"] textarea::placeholder {{
    color: rgba(255,255,255,0.45) !important;
  }}
  div[data-baseweb="input"]:focus-within, div[data-baseweb="textarea"]:focus-within {{
    border-color: {AZUL} !important; box-shadow: 0 0 0 3px rgba(24,95,165,0.30) !important;
  }}
  /* Menus flutuantes (opções do selectbox e o calendário do date_input) — no baseweb
     essas caixas têm fundo branco embutido em vários níveis, então zeramos tudo por
     dentro e pintamos só o nível de fora, escuro, translúcido. */
  div[data-baseweb="popover"] {{
    background: {AZUL_ESCURO} !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: 14px !important;
  }}
  div[data-baseweb="popover"] * {{ background: transparent !important; color: #ffffff !important; }}
  div[data-baseweb="popover"] li:hover,
  div[data-baseweb="popover"] [aria-selected="true"]:not([role="gridcell"]) {{
    background: rgba(255,255,255,0.14) !important;
  }}
  div[data-baseweb="calendar"] {{ background: {AZUL_ESCURO} !important; }}
  div[data-baseweb="calendar"] [role="gridcell"][aria-selected="true"] div {{
    background: {AZUL} !important; border-radius: 50% !important; color: #ffffff !important;
  }}
  div[data-baseweb="calendar"] [role="gridcell"]:hover div {{ background: rgba(255,255,255,0.16) !important; }}

  /* Botões em pílula, com feedback de pressão */
  .stButton > button, .stDownloadButton > button {{
    border-radius: 999px !important;
    transition: transform 140ms ease-out, background 160ms ease-out;
  }}
  .stButton > button:active, .stDownloadButton > button:active {{ transform: scale(0.97); }}
  .stButton > button[kind="primary"], .stDownloadButton > button {{
    background: rgba(255,255,255,0.94) !important; border: none !important; font-weight: 700 !important;
  }}
  .stButton > button[kind="primary"] p, .stDownloadButton > button p {{ color: {AZUL_ESCURO} !important; }}
  .stButton > button[kind="secondary"] {{
    background: rgba(255,255,255,0.08) !important; border: 1px solid rgba(255,255,255,0.24) !important;
  }}
  .stButton > button[kind="secondary"] p {{ color: #EAF3FC !important; }}

  /* Métricas como cards de vidro */
  div[data-testid="stMetric"] {{
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.16) !important;
    border-radius: 16px; padding: 16px 18px; backdrop-filter: blur(14px);
  }}
  div[data-testid="stMetric"] label {{ color: #AFCBE8 !important; font-weight: 600 !important; }}
  div[data-testid="stMetricValue"] {{ color: #ffffff !important; font-weight: 700; }}
  div[data-testid="stMetricDelta"] {{ color: #AFCBE8 !important; }}
</style>

<div class="topbar-vidro">
  <div class="marca"><span class="quad">SI</span> Soluções Industriais · Validador de Leads</div>
  <div class="saudacao">Olá, <b>{nome_analista}</b> 👋</div>
</div>

<div class="hero">
  <span class="badge">v3 · vidro industrial</span>
  <h1>Nova validação</h1>
  <p>Informe a chave única do cliente e o período — o app busca sozinho o briefing, os orçamentos e os anúncios no Metabase.</p>
</div>
""", unsafe_allow_html=True)

col_topo_vazio, col_trocar = st.columns([5, 1])
with col_trocar:
    if st.button("Trocar nome", use_container_width=True):
        st.session_state.pop("nome_analista", None)
        try:
            del st.query_params["analista"]
        except Exception:
            pass
        st.rerun()

CAMPOS_FORM = ("f_chave", "f_site", "f_obs", "f_modelo")
if st.session_state.pop("limpar_form", False):
    for k in CAMPOS_FORM:
        st.session_state.pop(k, None)
    st.session_state.pop("resultado", None)

with st.container(border=True):
    grid_a, grid_b = st.columns(2)
    with grid_a:
        chave_unica = st.text_input("Chave única do cliente", placeholder="Ex.: 12-34567-1", key="f_chave")
    with grid_b:
        st.markdown(
            "<div style='font-weight:600; font-size:0.82rem; color:#D9EAFB; margin-bottom:0.4rem;'>Período</div>",
            unsafe_allow_html=True,
        )
        p_a, p_b = st.columns(2)
        with p_a:
            data_inicio = st.date_input(
                "Início", value=date.today() - timedelta(days=90), format="DD/MM/YYYY", label_visibility="collapsed",
            )
        with p_b:
            data_fim = st.date_input(
                "Fim", value=date.today(), format="DD/MM/YYYY", label_visibility="collapsed",
            )

    grid_c, grid_d = st.columns(2)
    with grid_c:
        site = st.text_input(
            "Site do cliente (opcional, mas importante quando não há briefing cadastrado)",
            placeholder="https://www.sitedocliente.com.br", key="f_site",
        )
    with grid_d:
        obs = st.text_area(
            "Observações (opcional, mas importante quando não há briefing cadastrado)",
            placeholder="Ex.: cliente só vende máquinas (serviço, assistência, aluguel e peças = fora do foco).",
            height=90, key="f_obs",
        )

    linha_final_a, linha_final_b, linha_final_c = st.columns([2, 1, 1])
    with linha_final_a:
        modelo_escolha = st.selectbox(
            "Modelo de IA",
            list(MODELOS_ESCOLHA.keys()),
            key="f_modelo",
            help="Rápido = resposta mais veloz, boa para o dia a dia. Preciso = mais lento, "
                 "melhor em casos sutis. Automático usa o provedor com maior cota disponível.",
        )
    with linha_final_b:
        st.markdown("<div style='height:1.85rem;'></div>", unsafe_allow_html=True)
        if st.button("Limpar campos", use_container_width=True):
            st.session_state["limpar_form"] = True
            st.rerun()
    with linha_final_c:
        st.markdown("<div style='height:1.85rem;'></div>", unsafe_allow_html=True)
        validar = st.button("Validar leads", type="primary", use_container_width=True)


def montar_regras():
    if obs.strip():
        return f"- Observações do projeto (prioridade máxima): {obs.strip()}"
    return ""

if validar:
    ordem_ia = provedores_ativos()
    if not ordem_ia:
        st.error("Nenhuma chave de IA configurada. Adicione CEREBRAS_API_KEY ou GROQ_API_KEY nos Secrets.")
        st.stop()
    if not secret("METABASE_URL"):
        st.error("Segredo METABASE_URL não configurado (ex.: https://metabase.ferramentademarketing.com.br).")
        st.stop()
    if not chave_unica.strip():
        st.error("Preencha a chave única do cliente.")
        st.stop()

    # Traduz a escolha do seletor em ordem de provedores + modelo fixo (se houver)
    modelo_forcado = None
    escolha = MODELOS_ESCOLHA.get(modelo_escolha)
    if escolha:
        prov, modelo_forcado = escolha
        if prov not in ordem_ia:
            st.error(f"O modelo escolhido usa {prov.title()}, mas não há chave desse provedor nos Secrets. "
                     "Escolha outro modelo ou adicione a chave.")
            st.stop()
        ordem_ia = [prov] + [p for p in ordem_ia if p != prov]   # escolhido primeiro, resto como reserva

    nomes_ia = {"cerebras": "Cerebras", "groq": "Groq"}
    if modelo_forcado:
        st.caption(f"IA: {modelo_escolha}"
                   + (f" · reserva: {nomes_ia[ordem_ia[1]]}" if len(ordem_ia) > 1 else ""))
    else:
        st.caption("IA: " + " → ".join(nomes_ia[n] for n in ordem_ia)
                   + (" (reserva automática)" if len(ordem_ia) > 1 else ""))

    # 1. Briefing (question 286 — por chave única). Nem todo cliente tem briefing cadastrado.
    with st.spinner("Buscando briefing no Metabase..."):
        try:
            csv_briefing = consultar_question(CARD_BRIEFING, [
                ("chave_unica", chave_unica.strip(), "category"),
            ])
        except Exception as e:
            st.error(f"Erro ao buscar o briefing (question {CARD_BRIEFING}): {e}")
            st.stop()
    texto_briefing = csv_briefing_para_texto(csv_briefing)
    briefing_ausente = not texto_briefing
    if briefing_ausente:
        st.warning(
            f'Nenhum briefing cadastrado para a chave "{chave_unica}". '
            "A IA vai se basear no site informado e nas observações do projeto — "
            "preencha ao menos um desses dois campos para esse cliente."
        )

    nome_empresa = extrair_nome_empresa(csv_briefing) or chave_unica.strip()

    # 2. Orçamentos (question 47)
    with st.spinner("Buscando orçamentos no Metabase..."):
        try:
            csv_orcamentos = consultar_question(CARD_ORCAMENTOS, [
                ("chave_unica", chave_unica.strip(), "category"),
                ("data_inicio", data_inicio.isoformat(), "date/single"),
                ("data_fim", data_fim.isoformat(), "date/single"),
            ])
        except Exception as e:
            st.error(f"Erro ao buscar os orçamentos (question {CARD_ORCAMENTOS}): {e}")
            st.stop()
    linhas = list(csv.reader(io.StringIO(csv_orcamentos)))
    if len(linhas) < 2:
        st.error("Nenhum orçamento encontrado para essa chave única nesse período.")
        st.stop()

    # Se o briefing não trouxe o nome, tenta pelos orçamentos (coluna "Nome Fantasia")
    if nome_empresa == chave_unica.strip():
        nome_empresa = extrair_nome_empresa(csv_orcamentos) or nome_empresa
    st.caption(f"Cliente identificado: {nome_empresa}")

    # 3. Anúncios ativos (question 185 — opcional, não bloqueia se falhar)
    texto_anuncios = ""
    with st.spinner("Buscando anúncios do cliente no Metabase..."):
        try:
            csv_anuncios = consultar_question(CARD_ANUNCIOS, [
                ("chave_unica", chave_unica.strip(), "category"),
            ])
            texto_anuncios = csv_anuncios_para_texto(csv_anuncios)
        except Exception:
            st.warning("Não consegui buscar os anúncios (question 185) — prosseguindo sem eles.")

    # 4. Site do cliente (opcional — essencial quando não há briefing)
    texto_site = ""
    if site.strip():
        with st.spinner("Lendo o site do cliente..."):
            texto_site = buscar_site(site.strip())
        if not texto_site:
            st.warning("Não consegui ler o site informado — prosseguindo sem ele.")

    # 5. Perfil do cliente — monta com o que houver disponível (briefing e/ou site e/ou observações)
    partes_perfil = []
    regras_projeto = montar_regras()
    if regras_projeto:
        partes_perfil.append(f"===== OBSERVAÇÕES DO PROJETO (prioridade máxima) =====\n{regras_projeto}")
    if texto_briefing:
        partes_perfil.append(f"===== BRIEFING DO CLIENTE (Metabase) =====\n{texto_briefing}")
    if texto_site:
        partes_perfil.append(f"===== SITE DO CLIENTE ({site.strip()}) =====\n{texto_site}")
    if texto_anuncios:
        partes_perfil.append(f"===== ANÚNCIOS ATIVOS DO CLIENTE (termos anunciados) =====\n{texto_anuncios}")
    perfil = "\n\n".join(partes_perfil)

    if not perfil.strip():
        st.error(
            "Não há briefing, site nem observações suficientes para avaliar esse cliente. "
            "Preencha o site ou as observações do projeto e envie de novo."
        )
        st.stop()
    if len(perfil) > LIMITE_PERFIL:
        perfil = perfil[:LIMITE_PERFIL] + "\n[... perfil truncado para caber no limite da IA gratuita ...]"

    # 5. Montagem dos leads
    cabecalho = [c.strip() for c in linhas[0]]
    col_msg = "Mensagem do Cliente"
    if col_msg not in cabecalho:
        st.error(f'Coluna "{col_msg}" não encontrada no retorno do Metabase. Colunas: {", ".join(cabecalho)}')
        st.stop()
    idx_msg = cabecalho.index(col_msg)

    def idx_de(*nomes):
        for n in nomes:
            if n in cabecalho:
                return cabecalho.index(n)
        return -1

    idx_id = idx_de("ID do Orçamento", "ID do Orcamento")
    idx_nome = idx_de("Nome do Comprador", "Nome do comprador")
    idx_email = idx_de("E-mail do Comprador", "Email do Comprador", "E-mail do comprador")
    idx_anuncio = idx_de("anúncio de origem do Orçamento", "Anúncio do cliente", "anuncio de origem do Orçamento")

    registros = [r for r in linhas[1:] if any(c.strip() for c in r)]
    st.info(f"{len(registros)} leads encontrados de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}.")

    def celula(r, idx):
        return r[idx].strip() if 0 <= idx < len(r) else ""

    leads = []
    for i, r in enumerate(registros):
        extra = "; ".join(
            f"{cabecalho[j]}: {r[j]}" for j in range(len(cabecalho))
            if j != idx_msg and j < len(r) and r[j].strip()
        )[:500]
        leads.append({
            "id": f"L{i+2}",
            "mensagem": r[idx_msg] if idx_msg < len(r) else "",
            "extra": extra,
            "id_orc": celula(r, idx_id) or f"linha {i+2}",
            "nome": celula(r, idx_nome) or "(sem nome)",
            "email": celula(r, idx_email) or "(sem e-mail)",
            "anuncio": celula(r, idx_anuncio) or "(sem anúncio)",
            "tam_msg": len(celula(r, idx_msg)),
        })

    # 6. Classificação com re-tentativas
    classificacoes = {}
    erros_ia = []

    def processar(lista, tamanho_lote, rotulo):
        total_lotes = (len(lista) + tamanho_lote - 1) // tamanho_lote
        progresso = st.progress(0, text=f"{rotulo}: {len(lista)} leads...")
        for n in range(total_lotes):
            lote = lista[n * tamanho_lote:(n + 1) * tamanho_lote]
            try:
                resultado = chamar_ia(perfil, lote, ordem_ia, modelo_forcado)
            except RuntimeError as e:
                progresso.empty()
                st.error(str(e))       # cota diária esgotada: para tudo na hora
                st.stop()
            except Exception as e:
                erros_ia.append(f"{type(e).__name__}: {e}")
                resultado = []
            for item in resultado:
                status = str(item.get("status", "")).strip()
                if status not in STATUS_VALIDOS:
                    status = "Aberto"
                classificacoes[str(item.get("id", "")).strip()] = {
                    "status": status, "motivo": str(item.get("motivo", "")).strip(),
                }
            progresso.progress((n + 1) / total_lotes, text=f"{rotulo}: lote {n+1} de {total_lotes}")
            if n + 1 < total_lotes:
                time.sleep(1)
        progresso.empty()

    processar(leads, TAMANHO_LOTE, "Classificando")
    pendentes = [l for l in leads if l["id"] not in classificacoes]
    if pendentes:
        st.info(f"{len(pendentes)} lead(s) sem resposta — reprocessando em lotes menores...")
        time.sleep(5)
        processar(pendentes, 5, "Reprocessando")
    pendentes = [l for l in leads if l["id"] not in classificacoes]
    if pendentes:
        time.sleep(5)
        processar(pendentes, 1, "Última passada")
    falhas = sum(1 for l in leads if l["id"] not in classificacoes)

    # 7. Contagem + rankings para o dashboard
    contagem = {"Dentro do foco": 0, "Fora do foco": 0, "Aberto": 0}
    for i in range(len(registros)):
        c = classificacoes.get(leads[i]["id"])
        status = c["status"] if c else "Aberto"
        contagem[status] += 1
        leads[i]["status"] = status

    dentro = [l for l in leads if l.get("status") == "Dentro do foco"]
    fora = [l for l in leads if l.get("status") == "Fora do foco"]
    # melhores = dentro do foco com mensagem mais específica (mais longa) primeiro
    melhores = sorted(dentro, key=lambda l: l["tam_msg"], reverse=True)[:10]
    melhores = [{"id": l["id_orc"], "nome": l["nome"], "email": l["email"]} for l in melhores]
    piores = fora[:10]
    piores = [{"id": l["id_orc"], "nome": l["nome"], "email": l["email"]} for l in piores]
    # anúncios que mais geraram leads fora do foco
    cont_anuncios = {}
    for l in fora:
        cont_anuncios[l["anuncio"]] = cont_anuncios.get(l["anuncio"], 0) + 1
    anuncios_ruins = sorted(cont_anuncios.items(), key=lambda x: x[1], reverse=True)[:8]

    # 8. Excel com destaque + dashboard
    xlsx_bytes = gerar_xlsx(cabecalho, registros, leads, classificacoes)

    total = len(registros)
    periodo_txt = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
    base_nome = f"{nome_empresa} - {data_inicio.isoformat()} a {data_fim.isoformat()}"
    dash_html = gerar_dashboard_html(nome_empresa, chave_unica.strip(), periodo_txt, total, contagem,
                                     melhores=melhores, piores=piores, anuncios_ruins=anuncios_ruins)

    # Resultado fica guardado na sessão: os downloads não somem ao clicar
    st.session_state["resultado"] = {
        "empresa": nome_empresa,
        "total": total,
        "contagem": contagem,
        "falhas": falhas,
        "erro_ia": erros_ia[-1] if erros_ia else "",
        "xlsx_bytes": xlsx_bytes,
        "xlsx_nome": f"{base_nome} - Validado.xlsx",
        "dash_bytes": dash_html.encode("utf-8"),
        "dash_nome": f"{base_nome} - Dashboard.html",
    }

    salvar_no_historico({
        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
        "Data da solicitação": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Empresa": nome_empresa,
        "Chave única": chave_unica.strip(),
        "Período": periodo_txt,
        "Leads": total,
        "Dentro do foco": contagem["Dentro do foco"],
        "Fora do foco": contagem["Fora do foco"],
        "Aberto": contagem["Aberto"],
        "xlsx_nome": f"{base_nome} - Validado.xlsx",
        "dash_nome": f"{base_nome} - Dashboard.html",
    }, xlsx_bytes=xlsx_bytes,
       dash_bytes=st.session_state["resultado"]["dash_bytes"])

res = st.session_state.get("resultado")
if res:
    total, contagem = res["total"], res["contagem"]
    st.success(f"Validação concluída — {total} leads processados! ({res['empresa']})")
    c1, c2, c3 = st.columns(3)
    c1.metric("Dentro do foco", f"{contagem['Dentro do foco']/total:.0%}", f"{contagem['Dentro do foco']} leads", delta_color="off")
    c2.metric("Fora do foco", f"{contagem['Fora do foco']/total:.0%}", f"{contagem['Fora do foco']} leads", delta_color="off")
    c3.metric("Aberto", f"{contagem['Aberto']/total:.0%}", f"{contagem['Aberto']} leads", delta_color="off")
    if res["falhas"]:
        st.warning(f"{res['falhas']} lead(s) sem resposta da IA (marcados como Aberto) — rode de novo para reprocessar.")
        if res["erro_ia"]:
            st.error(f"Motivo técnico da falha na IA: {res['erro_ia'][:400]}")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button("Baixar Excel validado", data=res["xlsx_bytes"],
                           file_name=res["xlsx_nome"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, key="dl_xlsx")
    with col_dl2:
        st.download_button("Baixar dashboard (HTML)", data=res["dash_bytes"],
                           file_name=res["dash_nome"], mime="text/html",
                           use_container_width=True, key="dl_dash")

@st.dialog("Excluir pesquisa")
def dialogo_excluir(rid, rotulo):
    st.write(f"Tem certeza que deseja excluir a pesquisa **{rotulo}**?")
    st.caption("O registro e os arquivos (CSV e dashboard) dela serão apagados. Essa ação não tem volta.")
    cd1, cd2 = st.columns(2)
    if cd1.button("Sim, excluir", type="primary", use_container_width=True, key=f"conf_{rid}"):
        excluir_do_historico(rid)
        st.rerun()
    if cd2.button("Cancelar", use_container_width=True, key=f"canc_{rid}"):
        st.rerun()


st.markdown("<p style='font-weight:600; margin-top:32px;'>Histórico de validações</p>", unsafe_allow_html=True)
historico = carregar_historico()
if historico:
    filtro_chave = st.text_input(
        "Buscar por chave única",
        placeholder="Digite a chave para filtrar o histórico (ex.: 12-34567-1)",
        key="f_filtro_hist",
    )
    if filtro_chave.strip():
        alvo = filtro_chave.strip().lower()
        historico = [h for h in historico if alvo in str(h.get("Chave única", "")).lower()]
        if not historico:
            st.caption("Nenhuma validação encontrada para essa chave.")
    cab = st.columns([1.5, 1.6, 1.3, 1.9, 1.1, 0.4])
    for col, titulo in zip(cab, ("Data", "Empresa", "Chave", "Período", "Leads (D/F/A)", "")):
        col.markdown(f"<span style='font-size:0.72rem; color:#B5D4F4; font-weight:600;'>{titulo}</span>", unsafe_allow_html=True)

    for h in historico[:15]:
        rid = h.get("id", "")
        with st.container(border=True):
            c = st.columns([1.5, 1.6, 1.3, 1.9, 1.1, 0.4])
            c[0].markdown(f"<span style='font-size:0.78rem;'>{h.get('Data da solicitação', '')}</span>", unsafe_allow_html=True)
            c[1].markdown(f"<span style='font-size:0.78rem;'>{h.get('Empresa', '')}</span>", unsafe_allow_html=True)
            c[2].markdown(f"<span style='font-size:0.78rem;'>{h.get('Chave única', '')}</span>", unsafe_allow_html=True)
            c[3].markdown(f"<span style='font-size:0.78rem;'>{h.get('Período', '')}</span>", unsafe_allow_html=True)
            c[4].markdown(
                f"<span style='font-size:0.78rem;'>{h.get('Leads', '')} "
                f"(<span style='color:#22883A;'>{h.get('Dentro do foco', '')}</span>/"
                f"<span style='color:#D6433F;'>{h.get('Fora do foco', '')}</span>/"
                f"<span style='color:#B4830A;'>{h.get('Aberto', '')}</span>)</span>",
                unsafe_allow_html=True,
            )
            if rid and c[5].button("✕", key=f"x_{rid}", help="Excluir esta pesquisa"):
                dialogo_excluir(rid, f"{h.get('Empresa', '')} · {h.get('Data da solicitação', '')}")
    com_id = [h for h in historico if h.get("id")]
    if com_id:
        st.markdown("<p style='font-weight:600; margin-top:16px;'>Baixar arquivos de uma validação</p>", unsafe_allow_html=True)
        rotulos = {
            f"{h['Data da solicitação']} · {h.get('Empresa', '?')} · {h.get('Leads', '?')} leads": h
            for h in com_id[:15]
        }
        escolha = st.selectbox("Selecione a validação", list(rotulos.keys()), label_visibility="collapsed")
        sel = rotulos[escolha]
        rid = sel["id"]
        xlsx_salvo = ler_resultado_salvo(rid, ".xlsx")
        dash_salvo = ler_resultado_salvo(rid, ".html")
        cg1, cg2 = st.columns(2)
        with cg1:
            if xlsx_salvo:
                st.download_button("Baixar Excel", data=xlsx_salvo,
                                   file_name=sel.get("xlsx_nome", f"{rid}.xlsx"),
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, key=f"hxlsx_{rid}")
            else:
                st.caption("Excel não disponível (app reiniciou)")
        with cg2:
            if dash_salvo:
                st.download_button("Baixar dashboard", data=dash_salvo,
                                   file_name=sel.get("dash_nome", f"{rid}.html"),
                                   mime="text/html", use_container_width=True, key=f"hdash_{rid}")
            else:
                st.caption("Dashboard não disponível (app reiniciou)")
else:
    st.caption("Nenhuma validação registrada ainda. As próximas aparecerão aqui com data, empresa e resultado.")

st.markdown(
    "<p style='text-align:center; color:#B5D4F4; font-size:0.75rem; margin-top:32px;'>"
    "Validador de Leads v3 · Soluções Industriais</p>",
    unsafe_allow_html=True,
)
