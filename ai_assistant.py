"""Assistente de IA e extração de contexto para trabalhos acadêmicos do SCEN."""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

import docx
import pypdf

from gerar_trabalho import DADOS_ESSENCIAIS_TEMPLATE


def extrair_texto_arquivo(caminho: Path) -> str:
    """Extrai texto de arquivos PDF, DOCX, TXT ou MD."""
    sufixo = caminho.suffix.lower()
    if sufixo == ".pdf":
        reader = pypdf.PdfReader(str(caminho))
        paginas = [p.extract_text() or "" for p in reader.pages]
        return "\n".join(paginas).strip()
    if sufixo == ".docx":
        doc = docx.Document(str(caminho))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip()).strip()
    return caminho.read_text(encoding="utf-8", errors="replace").strip()


def montar_prompt_ia(
    tipo: str,
    dados_atuais: dict[str, Any],
    contexto_texto: str,
    arquivos_texto: list[tuple[str, str]],
    comando_usuario: str,
) -> str:
    """Monta o prompt estruturado pedindo retorno em JSON compatível com o formulário."""
    esquema = DADOS_ESSENCIAIS_TEMPLATE[tipo]
    partes_arquivos = []
    for nome, conteudo in arquivos_texto:
        trecho = conteudo[:12000]
        partes_arquivos.append(f"--- ARQUIVO ANEXADO: {nome} ---\n{trecho}")
    bloco_arquivos = "\n\n".join(partes_arquivos)

    return (
        "Você é um assistente acadêmico teológico do Seminário Cristão Evangélico do Norte (SCEN).\n"
        f"Tipo de trabalho selecionado: {tipo}\n\n"
        "OBJETIVO:\n"
        "Complete os campos do JSON abaixo com base no contexto, nos arquivos anexados e no comando do usuário.\n"
        "- Mantenha os campos já preenchidos pelo usuário, a menos que o comando peça para reescrever.\n"
        "- Responda SOMENTE com um objeto JSON válido (sem markdown, sem texto fora do JSON) seguindo exatamente as chaves do esquema.\n"
        "- Use português brasileiro acadêmico, claro e fiel às Escrituras e ao texto-base.\n\n"
        f"ESQUEMA DE CHAVES ESPERADO:\n{json.dumps(esquema, ensure_ascii=False, indent=2)}\n\n"
        f"DADOS ATUAIS DO FORMULÁRIO:\n{json.dumps(dados_atuais, ensure_ascii=False, indent=2)}\n\n"
        f"COMANDO / INSTRUÇÃO DO USUÁRIO:\n{comando_usuario or 'Completar todos os campos pendentes com base no contexto.'}\n\n"
        f"CONTEXTO E ANOTAÇÕES:\n{contexto_texto or '(nenhum texto colado)'}\n\n"
        f"{bloco_arquivos}"
    ).strip()


def extrair_json_resposta(texto: str) -> dict[str, Any]:
    """Extrai o objeto JSON da resposta da IA, mesmo se vier cercado por ```json."""
    limpo = texto.strip()
    m = re.search(r"```(?:json)?\s*(\{.*\})\s*```", limpo, flags=re.DOTALL)
    if m:
        limpo = m.group(1)
    else:
        inicio = limpo.find("{")
        fim = limpo.rfind("}")
        if inicio != -1 and fim != -1 and fim > inicio:
            limpo = limpo[inicio : fim + 1]
    return json.loads(limpo)


def _post_json(url: str, payload: dict[str, Any], headers: dict[str, str] | None = None) -> dict[str, Any]:
    data = json.dumps(payload).encode("utf-8")
    req_headers = {"Content-Type": "application/json"}
    if headers:
        req_headers.update(headers)
    req = urllib.request.Request(url, data=data, headers=req_headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detalhe = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Erro HTTP {exc.code}: {detalhe[:400]}") from exc


def chamar_ia(
    provedor: str,
    prompt: str,
    *,
    api_key: str = "",
    modelo: str = "",
    base_url: str = "",
) -> str:
    """Chama Gemini API, OpenAI/compatível ou Ollama local e retorna o texto gerado."""
    provedor = provedor.lower()

    if provedor == "gemini":
        if not api_key:
            raise ValueError("Informe a API Key do Gemini nas configurações.")
        mod = modelo or "gemini-2.5-pro"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent?key={api_key}"
        resp = _post_json(url, {"contents": [{"parts": [{"text": prompt}]}]})
        try:
            return resp["candidates"][0]["content"]["parts"][0]["text"]
        except (KeyError, IndexError) as exc:
            raise RuntimeError(f"Resposta inesperada do Gemini: {resp}") from exc

    if provedor == "openai":
        if not api_key:
            raise ValueError("Informe a API Key da OpenAI (ou provedor compatível).")
        mod = modelo or "gpt-4o-mini"
        endpoint = (base_url.rstrip("/") if base_url else "https://api.openai.com/v1") + "/chat/completions"
        resp = _post_json(
            endpoint,
            {
                "model": mod,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.4,
            },
            headers={"Authorization": f"Bearer {api_key}"},
        )
        try:
            return resp["choices"][0]["message"]["content"]
        except (KeyError, IndexError) as exc:
            raise RuntimeError(f"Resposta inesperada da OpenAI: {resp}") from exc

    if provedor == "ollama":
        mod = modelo or "llama3.1"
        endpoint = (base_url.rstrip("/") if base_url else "http://localhost:11434") + "/api/generate"
        resp = _post_json(endpoint, {"model": mod, "prompt": prompt, "stream": False})
        if "response" not in resp:
            raise RuntimeError(f"Resposta inesperada do Ollama: {resp}")
        return resp["response"]

    raise ValueError(f"Provedor desconhecido: {provedor}")


def listar_modelos_disponiveis(provedor: str, *, api_key: str = "", base_url: str = "") -> list[str]:
    """Consulta a API do provedor (Gemini, OpenAI ou Ollama) e retorna os modelos disponíveis."""
    prov = provedor.lower()
    if prov == "gemini":
        if not api_key:
            raise ValueError("Informe a API Key do Gemini para listar os modelos da sua conta.")
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        with urllib.request.urlopen(urllib.request.Request(url), timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        modelos = [
            m["name"].removeprefix("models/")
            for m in data.get("models", [])
            if "generateContent" in m.get("supportedGenerationMethods", [])
        ]
        return sorted(set(modelos), reverse=True)

    if prov == "openai":
        if not api_key:
            raise ValueError("Informe a API Key da OpenAI para listar os modelos.")
        endpoint = (base_url.rstrip("/") if base_url else "https://api.openai.com/v1") + "/models"
        req = urllib.request.Request(endpoint, headers={"Authorization": f"Bearer {api_key}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return sorted({m["id"] for m in data.get("data", []) if "id" in m}, reverse=True)

    if prov == "ollama":
        endpoint = (base_url.rstrip("/") if base_url else "http://localhost:11434") + "/api/tags"
        with urllib.request.urlopen(urllib.request.Request(endpoint), timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return [m["name"] for m in data.get("models", []) if "name" in m]

    raise ValueError(f"Provedor desconhecido: {provedor}")
