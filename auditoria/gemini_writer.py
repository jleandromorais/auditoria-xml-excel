import json
import os
import urllib.error
import urllib.request
from typing import Dict, List, Optional

import pandas as pd


def _try_load_dotenv():
    """Carrega .env se python-dotenv estiver instalado (silencioso)."""
    try:
        from pathlib import Path

        from dotenv import load_dotenv

        here = Path(__file__).resolve()
        # tenta alguns lugares comuns
        candidates = [
            here.parent / ".env",           # auditoria/.env
            here.parent.parent / ".env",    # auditoria-xml-excel/.env
            here.parent.parent.parent / ".env",  # raiz do workspace
        ]
        for p in candidates:
            if p.exists():
                load_dotenv(p)
    except Exception:
        return


def _as_float(v) -> Optional[float]:
    try:
        if v is None:
            return None
        return float(v)
    except Exception:
        return None


def _top_diffs(df: pd.DataFrame, col: str, n: int = 10) -> List[Dict]:
    if col not in df.columns:
        return []
    tmp = df.copy()
    tmp[col] = pd.to_numeric(tmp[col], errors="coerce")
    tmp = tmp.dropna(subset=[col])
    if tmp.empty:
        return []

    base_cols = [c for c in ["Nota", "Mes", "Empresa", "Tipo", "Status"] if c in tmp.columns]
    cols = base_cols + [col]

    # maiores positivos e maiores negativos
    pos = tmp.sort_values(col, ascending=False).head(n)[cols]
    neg = tmp.sort_values(col, ascending=True).head(n)[cols]

    out: List[Dict] = []
    for _, r in pos.iterrows():
        d = {k: (None if pd.isna(r[k]) else r[k]) for k in cols}
        d[col] = _as_float(d.get(col))
        d["_sentido"] = "POSITIVO"
        out.append(d)
    for _, r in neg.iterrows():
        d = {k: (None if pd.isna(r[k]) else r[k]) for k in cols}
        d[col] = _as_float(d.get(col))
        d["_sentido"] = "NEGATIVO"
        out.append(d)
    return out


def construir_payload_relatorio(df: pd.DataFrame) -> Dict:
    """Gera um payload enxuto (agregados + top exemplos) para a IA redigir o PDF."""
    status = df["Status"].astype(str) if "Status" in df.columns else pd.Series([], dtype=str)

    payload: Dict = {
        "metricas": {
            "total_analisado": int(len(df)),
            "ok": int(status.str.contains("OK", na=False).sum()) if not status.empty else 0,
            "erro": int(status.str.contains("ERRO", na=False).sum()) if not status.empty else 0,
            "sem_xml": int(status.str.contains("SEM XML", na=False).sum()) if not status.empty else 0,
            "sem_excel": int(status.str.contains("SEM EXCEL", na=False).sum()) if not status.empty else 0,
        },
        "top_diferencas": {
            "diff_rs": _top_diffs(df, "Diff R$", n=8),
            "diff_vol": _top_diffs(df, "Diff Vol", n=8),
        },
        "observacoes_exemplos": [],
    }

    # pega algumas observações reais (sem inventar)
    if "Obs" in df.columns:
        obs = (
            df["Obs"]
            .astype(str)
            .replace("nan", "")
            .replace("None", "")
        )
        obs = obs[obs.str.strip() != ""].head(10)
        payload["observacoes_exemplos"] = obs.tolist()

    return payload


def gerar_texto_pdf_com_gemini(
    df: pd.DataFrame,
    *,
    api_key: Optional[str] = None,
    # modelo atual da API (texto, multimodal) - v1
    model: str = "gemini-2.5-flash",
    timeout_s: int = 30,
) -> Optional[List[str]]:
    """
    Retorna uma lista de linhas para o PDF.
    Se não houver chave/erro, retorna None (caller faz fallback para texto fixo).
    """
    _try_load_dotenv()
    api_key = (api_key or os.getenv("GEMINI_API_KEY") or "").strip()
    if not api_key:
        print("[Gemini] GEMINI_API_KEY não encontrada (.env ou variável de ambiente). Usando texto padrão.")
        return None

    payload = construir_payload_relatorio(df)

    prompt = (
        "Você é um auditor fiscal/contábil experiente. Gere um texto para um PDF de 'Resumo Executivo' em português,\n"
        "explicando os resultados como se estivesse conversando com uma pessoa da área (analista fiscal/contador),\n"
        "e NÃO com um programador.\n"
        "REGRAS:\n"
        "- NÃO invente números nem fatos. Use SOMENTE os dados do JSON fornecido.\n"
        "- Se algum dado estiver ausente, escreva 'Não informado'.\n"
        "- Use linguagem clara, objetiva e profissional, com termos usados em auditoria fiscal/contábil.\n"
        "- Explique as causas prováveis das diferenças, impactos de risco fiscal e o que deve ser verificado na prática.\n"
        "- Estrutura obrigatória (com títulos):\n"
        "  1) Resumo Executivo\n"
        "  2) Escopo e Base de Dados\n"
        "  3) Principais Achados (com severidade: Alta/Média/Baixa)\n"
        "  4) Recomendações\n"
        "  5) Anexos (Top diferenças)\n"
        "- Produza no máximo 40 linhas curtas (para caber no PDF, como se fosse um parecer resumido).\n"
        "\n"
        f"DADOS (JSON):\n{json.dumps(payload, ensure_ascii=False)}\n"
    )

    # Endpoint v1 estável
    url = (
        "https://generativelanguage.googleapis.com/v1/"
        f"models/{model}:generateContent?key={api_key}"
    )

    body = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 900,
        },
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        try:
            detalhes = e.read().decode("utf-8", errors="replace")
            print(f"[Gemini] Erro HTTP {e.code}: {detalhes[:200]}")
        except Exception:
            print("[Gemini] Erro HTTP ao chamar a API.")
        return None
    except Exception as ex:
        print(f"[Gemini] Falha ao chamar a API: {ex}")
        return None

    try:
        data = json.loads(raw)
        text = (
            data["candidates"][0]["content"]["parts"][0]["text"]
            if data.get("candidates")
            else ""
        )
    except Exception as ex:
        print(f"[Gemini] Erro ao interpretar resposta: {ex}")
        return None

    linhas = [ln.rstrip() for ln in (text or "").splitlines()]
    linhas = [ln for ln in linhas if ln.strip() != ""]

    if not linhas:
        print("[Gemini] Resposta vazia. Usando texto padrão.")
        return None

    print("[Gemini] Texto gerado com sucesso para o PDF.")
    return linhas[:60]

