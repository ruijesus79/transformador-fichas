"""
Transformador de Fichas de Clientes
====================================
Aplicacao web para organizar ficheiros Excel/CSV de clientes desorganizados
em colunas limpas, prontas a importar no CallHub.

Abrir com duplo clique no ficheiro 'Abrir Transformador.bat'
ou executar: python app.py
"""

import functools
import gc
import io
import os
import re
import shutil
import threading
import time
import unicodedata
import uuid
import webbrowser
import zipfile
from datetime import date, datetime

import pandas as pd
from flask import Flask, abort, jsonify, request, send_from_directory

app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Limite de 50 MB por pedido
EXTENSOES_PERMITIDAS = {'xlsx', 'xls', 'csv'}
MINUTOS_VALIDADE = 15  # ficheiros gerados sao apagados ao fim deste tempo

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_ficheiros_gerados')
os.makedirs(OUTPUT_DIR, exist_ok=True)


class ErroUtilizador(Exception):
    """Erro com mensagem que pode ser mostrada diretamente ao utilizador."""


# ─── HTML Template ────────────────────────────────────────────────────────────

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Transformador de Fichas de Clientes</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: #fff;
            padding: 20px 0;
        }

        .container { max-width: 760px; width: 92%; text-align: center; }

        .logo { font-size: 3rem; margin-bottom: 0.5rem; animation: pulse 2s ease-in-out infinite; }
        @keyframes pulse { 0%, 100% { transform: scale(1); } 50% { transform: scale(1.05); } }

        h1 { font-size: 1.8rem; font-weight: 300; margin-bottom: 0.3rem; letter-spacing: 1px; }
        .subtitle { color: #a0a0c0; font-size: 0.95rem; margin-bottom: 2rem; }

        .drop-zone {
            border: 3px dashed rgba(255, 255, 255, 0.25);
            border-radius: 20px;
            padding: 60px 40px;
            transition: all 0.3s ease;
            cursor: pointer;
            background: rgba(255, 255, 255, 0.03);
        }
        .drop-zone:hover, .drop-zone.dragover {
            border-color: #7c5cfc;
            background: rgba(124, 92, 252, 0.08);
            transform: scale(1.02);
        }
        .drop-zone-icon { font-size: 4rem; margin-bottom: 1rem; opacity: 0.7; }
        .drop-zone-text { font-size: 1.2rem; color: #c0c0e0; margin-bottom: 0.5rem; }
        .drop-zone-hint { font-size: 0.85rem; color: #808090; }
        .drop-zone input[type="file"] { display: none; }

        .options {
            margin-top: 18px;
            display: flex;
            flex-direction: column;
            gap: 8px;
            align-items: center;
        }
        .options label {
            cursor: pointer;
            font-size: 0.95rem;
            color: #a0a0c0;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .options input { width: 16px; height: 16px; cursor: pointer; }
        .options small { color: #707088; }

        .processing { display: none; padding: 60px 40px; }
        .spinner {
            width: 60px; height: 60px;
            border: 4px solid rgba(255, 255, 255, 0.1);
            border-left-color: #7c5cfc;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin: 0 auto 1.5rem;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .processing-text { font-size: 1.1rem; color: #c0c0e0; }
        .processing-hint { font-size: 0.85rem; color: #808090; margin-top: 0.5rem; }

        .result { display: none; padding: 30px 0; }
        .result-success {
            background: rgba(76, 175, 80, 0.1);
            border: 2px solid rgba(76, 175, 80, 0.3);
            border-radius: 20px;
            padding: 30px;
        }
        .result-icon { font-size: 3.5rem; margin-bottom: 0.8rem; }
        .result-title { font-size: 1.4rem; font-weight: 600; margin-bottom: 0.5rem; color: #4caf50; }
        .result-stats { display: flex; justify-content: center; gap: 1.6rem; margin: 1.5rem 0; flex-wrap: wrap; }
        .stat { text-align: center; min-width: 90px; }
        .stat-value { font-size: 1.8rem; font-weight: 700; color: #7c5cfc; }
        .stat-label { font-size: 0.72rem; color: #808090; text-transform: uppercase; letter-spacing: 1px; }

        .btn {
            display: inline-block;
            padding: 14px 32px;
            border-radius: 50px;
            font-size: 1rem;
            font-weight: 600;
            text-decoration: none;
            transition: all 0.3s ease;
            cursor: pointer;
            border: none;
            margin: 0.4rem;
        }
        .btn-download {
            background: linear-gradient(135deg, #7c5cfc, #5a3fd4);
            color: #fff;
            box-shadow: 0 4px 20px rgba(124, 92, 252, 0.4);
        }
        .btn-download:hover { transform: translateY(-2px); box-shadow: 0 6px 30px rgba(124, 92, 252, 0.6); }
        .btn-callhub { background: linear-gradient(135deg, #20a36b, #138a57); box-shadow: 0 4px 20px rgba(32, 163, 107, 0.4); }
        .btn-reset { background: rgba(255, 255, 255, 0.1); color: #c0c0e0; }
        .btn-reset:hover { background: rgba(255, 255, 255, 0.15); }

        .note { font-size: 0.8rem; color: #8a8aa0; margin-top: 1rem; line-height: 1.5; }

        .result-error {
            background: rgba(244, 67, 54, 0.1);
            border: 2px solid rgba(244, 67, 54, 0.3);
            border-radius: 20px;
            padding: 40px;
        }
        .result-error .result-title { color: #f44336; }
        .error-message { color: #e0a0a0; font-size: 0.95rem; margin-top: 1rem; word-break: break-word; }

        .footer { margin-top: 2rem; color: #505060; font-size: 0.75rem; }

        .preview-table { margin: 1.5rem auto; max-width: 100%; overflow-x: auto; }
        .preview-table table { border-collapse: collapse; font-size: 0.75rem; width: 100%; }
        .preview-table th {
            background: rgba(124, 92, 252, 0.2);
            color: #c0c0e0;
            padding: 8px 12px;
            text-align: left;
            white-space: nowrap;
        }
        .preview-table td {
            padding: 6px 12px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #a0a0b0;
            white-space: nowrap;
            max-width: 260px;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .preview-table tr:hover td { background: rgba(255,255,255,0.03); }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">&#128203;</div>
        <h1>Transformador de Fichas</h1>
        <p class="subtitle">Arrasta o ficheiro e os dados ficam organizados por colunas, prontos para o CallHub</p>

        <div id="uploadArea">
            <div class="drop-zone" id="dropZone">
                <div class="drop-zone-icon">&#128194;</div>
                <div class="drop-zone-text">Arrasta o(s) ficheiro(s) para aqui</div>
                <div class="drop-zone-hint">ou clica para selecionar (.xlsx, .xls, .csv &middot; m&aacute;x. 50 MB)</div>
                <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" multiple>
            </div>

            <div class="options">
                <label>
                    <input type="checkbox" id="mergeCheckbox" checked>
                    Juntar todos os ficheiros num &uacute;nico ficheiro final (recomendado)
                </label>
                <label>
                    <input type="checkbox" id="prefixCheckbox">
                    Adicionar indicativo +351 aos contactos <small>(se o CallHub o exigir)</small>
                </label>
            </div>
        </div>

        <div class="processing" id="processing">
            <div class="spinner"></div>
            <div class="processing-text">A organizar ficheiro(s)...</div>
            <div class="processing-hint">Ficheiros grandes podem demorar at&eacute; 1 minuto. N&atilde;o feches esta p&aacute;gina.</div>
        </div>

        <div class="result" id="result"></div>

        <div class="footer">
            Equipa Red Team Alcino Fontes &middot; Transformador de Fichas v3.0
        </div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        const processing = document.getElementById('processing');
        const result = document.getElementById('result');
        const mergeCheckbox = document.getElementById('mergeCheckbox');
        const prefixCheckbox = document.getElementById('prefixCheckbox');

        // Lembrar as opcoes escolhidas (merge ativo por defeito)
        [['opt_merge', mergeCheckbox, true], ['opt_prefix', prefixCheckbox, false]].forEach(([key, el, defVal]) => {
            try {
                const salvo = localStorage.getItem(key);
                el.checked = salvo !== null ? salvo === '1' : defVal;
            } catch (e) {
                el.checked = defVal;
            }
            el.addEventListener('change', () => {
                try { localStorage.setItem(key, el.checked ? '1' : '0'); } catch (e) {}
            });
        });

        function esc(value) {
            return String(value === null || value === undefined ? '' : value)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
        }

        function mensagemPorEstado(status) {
            if (status === 413) return 'O ficheiro é demasiado grande (máximo 50 MB).';
            if (status === 502 || status === 503 || status === 504) {
                return 'O servidor demorou demasiado a responder. Se o site esteve parado, aguarda 1 minuto e tenta novamente.';
            }
            return 'Ocorreu um erro inesperado (código ' + status + '). Tenta novamente.';
        }

        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) processFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) processFiles(fileInput.files);
        });

        let ultimosFicheiros = null;

        function processFiles(files, tentativa = 1) {
            const validFiles = Array.from(files).filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
            if (validFiles.length === 0) {
                showError('Seleciona ficheiro(s) Excel (.xlsx, .xls) ou CSV.');
                return;
            }
            ultimosFicheiros = validFiles;

            uploadArea.style.display = 'none';
            processing.style.display = 'block';
            result.style.display = 'none';

            const procText = document.querySelector('.processing-text');
            const procHint = document.querySelector('.processing-hint');
            if (tentativa > 1) {
                if (procText) procText.textContent = `A ligar ao servidor (tentativa ${tentativa} de 3)...`;
                if (procHint) procHint.textContent = 'O servidor gratuito no Render entra em repouso após inatividade e demora ~30s a acordar.';
            } else {
                if (procText) procText.textContent = 'A organizar ficheiro(s)...';
                if (procHint) procHint.textContent = 'Ficheiros grandes podem demorar até 1 minuto. Não feches esta página.';
            }

            const formData = new FormData();
            validFiles.forEach(f => formData.append('file', f));
            if (mergeCheckbox.checked) formData.append('merge', 'true');
            if (prefixCheckbox.checked) formData.append('prefixo351', 'true');

            fetch('/transformar', { method: 'POST', body: formData })
                .then(async (response) => {
                    let data = null;
                    try { data = await response.json(); } catch (e) { data = null; }
                    if (!response.ok || !data) {
                        throw new Error((data && data.error) || mensagemPorEstado(response.status));
                    }
                    return data;
                })
                .then(data => showSuccess(data, validFiles.length))
                .catch(error => {
                    const isNetwork = !error || !error.message || error.message === 'Failed to fetch' || error.message.includes('NetworkError');
                    if (isNetwork && tentativa < 3) {
                        setTimeout(() => processFiles(validFiles, tentativa + 1), 3500);
                        return;
                    }
                    const msg = (!isNetwork)
                        ? error.message
                        : 'O servidor na nuvem estava a acordar ou reiniciou. Clica em "Tentar novamente" abaixo para concluir.';
                    showError(msg);
                });
        }

        function showSuccess(data, numFiles) {
            processing.style.display = 'none';
            result.style.display = 'block';

            let previewHtml = '';
            if (data.preview && data.preview.length > 0) {
                const cols = Object.keys(data.preview[0]);
                previewHtml = `
                    <div class="preview-table">
                        <table>
                            <thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead>
                            <tbody>${data.preview.map(row =>
                                `<tr>${cols.map(c => `<td title="${esc(row[c])}">${esc(row[c])}</td>`).join('')}</tr>`
                            ).join('')}</tbody>
                        </table>
                    </div>`;
            }

            const s = data.stats;
            const title = numFiles > 1 ? `${numFiles} ficheiros organizados!` : 'Ficheiro organizado!';
            const botoes = data.downloads.map(d =>
                `<a href="${esc(d.url)}" class="btn btn-download ${d.tipo === 'csv' ? 'btn-callhub' : ''}">&#11015; ${esc(d.label)}</a>`
            ).join('');

            result.innerHTML = `
                <div class="result-success">
                    <div class="result-icon">&#9989;</div>
                    <div class="result-title">${esc(title)}</div>
                    <div class="result-stats">
                        <div class="stat"><div class="stat-value">${s.linhas_lidas}</div><div class="stat-label">Linhas lidas</div></div>
                        <div class="stat"><div class="stat-value">${s.com_contacto}</div><div class="stat-label">Com contacto</div></div>
                        <div class="stat"><div class="stat-value">${s.multi_contacto}</div><div class="stat-label">2+ contactos</div></div>
                        <div class="stat"><div class="stat-value">${s.sem_contacto}</div><div class="stat-label">Sem contacto</div></div>
                        <div class="stat"><div class="stat-value">${s.duplicados}</div><div class="stat-label">Duplicados removidos</div></div>
                    </div>
                    ${previewHtml}
                    <div>${botoes}</div>
                    <button class="btn btn-reset" onclick="resetForm()">&#128260; Novo(s) ficheiro(s)</button>
                    <div class="note">
                        O ficheiro CSV (verde) contém apenas os clientes com contacto e está pronto a importar no CallHub.<br>
                        O Excel tem a folha "Clientes" e, se existirem, a folha "Sem Contacto".<br>
                        Por privacidade, os ficheiros são apagados do servidor ao fim de ${data.validade_minutos} minutos.
                    </div>
                </div>`;
        }

        function showError(message) {
            processing.style.display = 'none';
            result.style.display = 'block';
            result.innerHTML = `
                <div class="result-error">
                    <div class="result-icon">&#10060;</div>
                    <div class="result-title">Não foi possível processar</div>
                    <div class="error-message">${esc(message)}</div>
                    <br>
                    <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
                        <button class="btn btn-download" style="margin: 0;" onclick="tentarNovamente()">&#128260; Tentar novamente</button>
                        <button class="btn btn-reset" style="margin: 0; background: rgba(255,255,255,0.1);" onclick="resetForm()">&#128194; Escolher outro</button>
                    </div>
                </div>`;
        }

        function tentarNovamente() {
            if (ultimosFicheiros && ultimosFicheiros.length > 0) {
                processFiles(ultimosFicheiros);
            } else {
                resetForm();
            }
        }

        function resetForm() {
            uploadArea.style.display = 'block';
            processing.style.display = 'none';
            result.style.display = 'none';
            fileInput.value = '';
            ultimosFicheiros = null;
        }
    </script>
</body>
</html>
"""


# ─── Expressoes regulares e constantes ─────────────────────────────────────────

RE_MEO = re.compile(r'\b(NIC|DIF)\s*:', re.I)
RE_ROTULOS_ID = re.compile(
    r'\b(NIC|DIF|DIP|NIF|NIPC|CONTRIBUINTE|CC|BI)\s*[:=]\s*[\w\-/]*', re.I)
RE_NIF_ROTULO = re.compile(r'\b(?:NIF|NIPC|DIF|CONTRIBUINTE)\s*[:.=]?\s*(\d{9})\b', re.I)
RE_CANDIDATO_TEL = re.compile(r'(?:\+|00)?\d[\d\s.\-()/]{7,}\d')
RE_EMAIL = re.compile(r'[\w.+-]+@[\w-]+\.[\w.-]+')
RE_CP = re.compile(r'\b(\d{4})\s?-\s?(\d{3})\b')
RE_DATA_TEXTO = re.compile(r'^\s*(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})(\s+[\d:]+)?\s*$')
RE_RUA = re.compile(
    r'^(rua|r\.|av\.?|avenida|travessa|tv\.?|trav\.?|largo|lg\.?|praca|praceta|pc\.?|estrada|'
    r'est\.?|beco|caminho|alameda|urbanizacao|urb\.?|quinta|lugar|bairro|calcada|rotunda|'
    r'azinhaga|rampa|loteamento|lote|edificio|zona)\b')
RE_EMPRESA = re.compile(
    r'(?<!\w)(LDA|L\.DA|UNIPESSOAL|S\.A\.?|SOCIEDADE|LIMITADA|ASSOCIACAO|CLUBE|RESTAURANTE|'
    r'CAFE|OFICINA|LOJA|HOTEL|HOTELS|CONDOMINIO|JUNTA DE FREGUESIA|CAMARA MUNICIPAL|FUNDACAO|'
    r'COOPERATIVA|COMPANHIA|SA)(?!\w)')
RE_SERVICO = re.compile(
    r'\b(FIBRA|ADSL|SAT|TV\+NET|TV\+VOZ|NET\+VOZ|MEO\s*BOX|MEO\s*GO|M[345]O|TARIF|PACOTE|TELEMOVEL|MOVEL)\b',
    re.I)
PARTICULAS = {'da', 'de', 'do', 'das', 'dos', 'e', 'd'}
VALORES_VAZIOS = {'', 'nan', 'none', 'nat', '<na>', 'null'}

# Papel de cada coluna a partir do texto do cabecalho (ordem importa)
CABECALHOS = [
    ('nic', ['nic', 'n cliente', 'no cliente', 'num cliente', 'numero cliente', 'numero de cliente',
             'n de cliente', 'conta cliente', 'id cliente', 'n conta']),
    ('nif', ['nif', 'nipc', 'contribuinte', 'dif']),
    ('email', ['email', 'e mail', 'mail']),
    ('telefone', ['telefone', 'telemovel', 'telm', 'tlm', 'tlf', 'tel', 'contacto', 'contato',
                  'contactos', 'contatos', 'movel', 'phone', 'fixo']),
    ('cp', ['codigo postal', 'cod postal', 'cp', 'cp7', 'postal']),
    ('morada', ['morada', 'endereco', 'rua', 'address', 'residencia']),
    ('localidade', ['localidade', 'cidade', 'concelho', 'freguesia', 'distrito', 'local']),
    ('data', ['fidelizacao', 'data', 'fim fidelizacao', 'validade']),
    ('operadora', ['operador', 'operadora', 'fornecedor']),
    ('servico', ['servico', 'tarifario', 'pacote', 'produto', 'plano']),
    ('nome', ['nome', 'cliente', 'titular', 'name']),
]
PAPEIS_UNICOS = {'nome', 'nif', 'nic', 'email', 'cp', 'morada', 'localidade', 'data', 'operadora', 'servico'}


# ─── Funcoes auxiliares ────────────────────────────────────────────────────────

@functools.lru_cache(maxsize=32768)
def sem_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))


def celula_texto(valor):
    """Converte qualquer celula em texto limpo (sem '.0' nos numeros, sem 'nan')."""
    if valor is None:
        return ''
    if isinstance(valor, float):
        if pd.isna(valor):
            return ''
        return str(int(valor)) if valor.is_integer() else str(valor)
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        if pd.isna(valor):
            return ''
        return valor.strftime('%d/%m/%Y')
    texto = str(valor).strip()
    return '' if texto.lower() in VALORES_VAZIOS else texto


def letra_coluna(indice):
    letras = ''
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def nif_valido(valor):
    s = re.sub(r'\D', '', str(valor))
    if len(s) != 9 or s[0] == '0':
        return False
    total = sum(int(s[i]) * (9 - i) for i in range(8))
    resto = total % 11
    controlo = 0 if resto < 2 else 11 - resto
    return controlo == int(s[8])


def _digitos_para_telefones(d):
    """Recebe so digitos e devolve a lista de telefones PT (9 digitos) que representam."""
    if len(d) == 9 and d[0] in '29':
        return [d]
    if len(d) == 12 and d.startswith('351') and d[3] in '29':
        return [d[3:]]
    if len(d) == 14 and d.startswith('00351') and d[5] in '29':
        return [d[5:]]
    if len(d) > 9 and len(d) % 9 == 0:
        partes = [d[i:i + 9] for i in range(0, len(d), 9)]
        if all(p[0] in '29' for p in partes):
            return partes
    return []


def _telefones_de_candidato(candidato):
    encontrados = _digitos_para_telefones(re.sub(r'\D', '', candidato))
    if encontrados:
        return encontrados
    # Tenta separar por espacos (ex: "916380107 30" -> 916380107)
    tokens = candidato.split()
    resultado, i = [], 0
    while i < len(tokens):
        for j in range(len(tokens), i, -1):
            achados = _digitos_para_telefones(re.sub(r'\D', '', ''.join(tokens[i:j])))
            if achados:
                resultado.extend(achados)
                i = j
                break
        else:
            i += 1
    return resultado


def extrair_telefones(texto):
    """Extrai telefones portugueses de qualquer texto (aceita +351, 00351, espacos, hifenes)."""
    if not texto:
        return []
    limpo = RE_ROTULOS_ID.sub(' ', str(texto))
    limpo = RE_EMAIL.sub(' ', limpo)
    telefones = []
    for candidato in RE_CANDIDATO_TEL.findall(limpo):
        for tel in _telefones_de_candidato(candidato):
            if tel not in telefones:
                telefones.append(tel)
    return telefones


def normalizar_cabecalho(texto):
    t = sem_acentos(str(texto)).lower()
    t = re.sub(r'[^a-z0-9]+', ' ', t)
    t = re.sub(r'\b\d+\b', ' ', t)  # "Telefone 2" -> "telefone"
    return re.sub(r'\s+', ' ', t).strip()


def papel_do_cabecalho(texto):
    h = normalizar_cabecalho(texto)
    if not h:
        return None
    tokens = set(h.split())
    for papel, chaves in CABECALHOS:
        for chave in chaves:
            if h == chave or chave in tokens or (len(chave) > 4 and chave in h):
                return papel
    return None


def converter_data(valor):
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        ts = pd.Timestamp(valor)
        return None if pd.isna(ts) else ts.normalize()
    texto = celula_texto(valor)
    if not texto or not RE_DATA_TEXTO.match(texto):
        return None
    iso = re.match(r'^\s*\d{4}-', texto) is not None
    ts = pd.to_datetime(texto, dayfirst=not iso, errors='coerce')
    return None if pd.isna(ts) else ts.normalize()


def limpar_nome(nome):
    if not nome:
        return ''
    t = re.sub(r'^\s*(nome|cliente|titular)\s*:\s*', '', str(nome), flags=re.I)
    return re.sub(r'\s+', ' ', t).strip(' ,;-')


def e_empresa(nome):
    return bool(nome) and RE_EMPRESA.search(sem_acentos(nome).upper()) is not None


def nome_em_titulo(nome):
    palavras = []
    for i, palavra in enumerate(nome.lower().split()):
        if i > 0 and palavra in PARTICULAS:
            palavras.append(palavra)
        else:
            palavras.append('-'.join(p[:1].upper() + p[1:] for p in palavra.split('-')))
    return ' '.join(palavras)


def tipo_cliente(nif, nome):
    if nome and e_empresa(nome):
        return 'Empresa'
    if nif:
        return 'Empresa' if str(nif)[0] in '5679' else 'Particular'
    return 'Particular' if nome else ''


# ─── Leitura de ficheiros ──────────────────────────────────────────────────────

def ler_ficheiro(ficheiro):
    """Devolve lista de (nome_folha, DataFrame) com TODAS as folhas do ficheiro."""
    nome = ficheiro.filename
    dados = ficheiro.read()
    if not dados:
        raise ErroUtilizador(f'O ficheiro "{nome}" está vazio.')
    extensao = nome.rsplit('.', 1)[-1].lower()
    try:
        if extensao == 'csv':
            for codificacao in ('utf-8-sig', 'cp1252', 'latin-1'):
                try:
                    df = pd.read_csv(io.BytesIO(dados), header=None, sep=None, engine='python',
                                     encoding=codificacao, dtype=object)
                    return [('CSV', df)]
                except UnicodeDecodeError:
                    continue
            raise ErroUtilizador(f'Não foi possível ler o CSV "{nome}" (codificação desconhecida).')
        if extensao == 'xlsx':
            wb = openpyxl.load_workbook(io.BytesIO(dados), read_only=True, data_only=True)
            folhas = []
            for sname in wb.sheetnames:
                sheet = wb[sname]
                rows = list(sheet.iter_rows(values_only=True))
                if rows:
                    folhas.append((sname, pd.DataFrame(rows)))
            wb.close()
            if not folhas:
                raise ErroUtilizador(f'O ficheiro "{nome}" não contém dados.')
            return folhas
        motor = 'xlrd' if extensao == 'xls' else 'openpyxl'
        folhas = pd.read_excel(io.BytesIO(dados), header=None, sheet_name=None, engine=motor)
        return list(folhas.items())
    except ErroUtilizador:
        raise
    except Exception as exc:
        app.logger.warning('Falha a ler %s: %s', nome, exc)
        raise ErroUtilizador(
            f'Não foi possível abrir "{nome}". Confirma que é um ficheiro Excel/CSV válido '
            f'e que não está protegido por palavra-passe.')


# ─── Detecao do papel de cada coluna ───────────────────────────────────────────

def _fracao(textos, condicao):
    return sum(1 for t in textos if condicao(t)) / len(textos) if textos else 0


def perfil_coluna(valores):
    brutos = [v for v in valores if celula_texto(v)][:400]
    if not brutos:
        return 'vazia'
    textos = [celula_texto(v) for v in brutos]
    n = len(textos)

    if _fracao(textos, lambda t: RE_MEO.search(t)) > 0.3:
        return 'meo'
    if sum(1 for v in brutos if converter_data(v) is not None) / n > 0.6:
        return 'data'
    if _fracao(textos, lambda t: RE_CP.fullmatch(t.strip())) > 0.6:
        return 'cp'
    if _fracao(textos, lambda t: RE_EMAIL.search(t)) > 0.6:
        return 'email'
    if _fracao(textos, lambda t: not re.search(r'[A-Za-z]', t) and nif_valido(t)
               and len(re.sub(r'\D', '', t)) == 9) > 0.7:
        return 'nif'
    if _fracao(textos, lambda t: extrair_telefones(t) and len(re.findall(r'[A-Za-z]', t)) <= 4) > 0.6:
        return 'telefone'
    if _fracao(textos, lambda t: RE_SERVICO.search(t)) > 0.35:
        return 'servico'
    if _fracao(textos, lambda t: RE_CP.search(t) or RE_RUA.match(sem_acentos(t).lower())) > 0.5:
        return 'morada'
    if _fracao(textos, lambda t: extrair_telefones(t)) > 0.3:
        return 'misto'
    if _fracao(textos, lambda t: re.search(r'[A-Za-zÀ-ÿ]', t) and not re.search(r'\d', t)) > 0.7:
        return 'texto'
    return 'outros'


def detetar_cabecalho(df):
    """Procura uma linha de cabecalho nas primeiras 5 linhas. Devolve (indice, nomes) ou (None, None)."""
    for i in range(min(5, len(df))):
        celulas = [celula_texto(v) for v in df.iloc[i]]
        preenchidas = [c for c in celulas if c]
        if not preenchidas:
            continue
        if any(RE_MEO.search(c) or extrair_telefones(c) for c in preenchidas):
            return None, None
        reconhecidas = sum(1 for c in preenchidas if len(c) <= 40 and papel_do_cabecalho(c))
        minimo = 1 if len(preenchidas) <= 2 else 2
        if reconhecidas >= minimo and reconhecidas >= len(preenchidas) * 0.4:
            return i, celulas
    return None, None


def decidir_papeis(df, cabecalho):
    papeis = {}
    for col in df.columns:
        perfil = perfil_coluna(df[col].tolist())
        papel_cab = papel_do_cabecalho(cabecalho[col]) if cabecalho and cabecalho[col] else None
        if perfil in ('meo', 'vazia'):
            papeis[col] = perfil
        elif papel_cab:
            papeis[col] = papel_cab
        else:
            papeis[col] = perfil

    # Resolver colunas de texto simples: nome vs localidade
    colunas_texto = [c for c, p in papeis.items() if p == 'texto']
    if colunas_texto:
        col_morada = next((c for c, p in papeis.items() if p == 'morada'), None)
        ja_tem_nome = any(p in ('nome', 'meo') for p in papeis.values())
        metricas = {}
        for c in colunas_texto:
            amostra = df[[c] + ([col_morada] if col_morada is not None else [])].head(400)
            textos = [celula_texto(v) for v in amostra[c] if celula_texto(v)]
            contido = 0.0
            if col_morada is not None and textos:
                pares = [(celula_texto(a).lower(), celula_texto(b).lower())
                         for a, b in zip(amostra[c], amostra[col_morada]) if celula_texto(a)]
                contido = sum(1 for a, b in pares if a and a in b) / len(pares)
            unicidade = len(set(textos)) / len(textos) if textos else 0
            palavras = sum(len(t.split()) for t in textos) / len(textos) if textos else 0
            metricas[c] = (contido, unicidade, palavras)

        tem_localidade = any(p == 'localidade' for p in papeis.values())
        for c in sorted(colunas_texto, key=lambda x: -metricas[x][0]):
            contido, unicidade, palavras = metricas[c]
            if not tem_localidade and contido > 0.5:
                papeis[c] = 'localidade'
                tem_localidade = True
        for c in sorted(colunas_texto, key=lambda x: -(metricas[x][1] * metricas[x][2])):
            if papeis[c] != 'texto':
                continue
            contido, unicidade, palavras = metricas[c]
            if not ja_tem_nome and (palavras >= 1.5 or unicidade >= 0.7):
                papeis[c] = 'nome'
                ja_tem_nome = True
            elif not tem_localidade and unicidade < 0.5:
                papeis[c] = 'localidade'
                tem_localidade = True
            else:
                papeis[c] = 'outros'
    return papeis


# ─── Interpretacao de cada linha ───────────────────────────────────────────────

def registo_vazio(origem):
    return {'nome': '', 'telefones': [], 'nif': '', 'nic': '', 'operadora': '', 'servico': '',
            'data': None, 'morada': '', 'localidade': '', 'cp': '', 'email': '', 'outros': [],
            'origem': origem}


def interpretar_meo(texto, reg):
    """Formato do CRM da operadora: 'NOME; NIC: ..; DIF: ..; DIP: ..; SERVICO; tel1,tel2'."""
    partes = [p.strip() for p in texto.split(';')]
    if partes and not reg['nome']:
        reg['nome'] = limpar_nome(partes[0])

    m = re.search(r'\bNIC\s*:\s*(\d+)', texto, re.I)
    if m and not reg['nic']:
        reg['nic'] = m.group(1)
    m = re.search(r'\bDIF\s*:\s*(\d+)', texto, re.I)
    if m and not reg['nif']:
        reg['nif'] = m.group(1)

    servicos, telefones = [], []
    for parte in partes[1:]:
        conteudo = RE_ROTULOS_ID.sub(' ', parte).strip(' ,\n\r\t')
        if not conteudo:
            continue
        if re.search(r'[A-Za-z]', conteudo):
            servicos.append(re.sub(r'\s+', ' ', conteudo))
        else:
            telefones.extend(extrair_telefones(conteudo))
    if not telefones:  # recurso: procurar no texto todo, excepto no servico
        texto_sem_servico = texto
        for s in servicos:
            texto_sem_servico = texto_sem_servico.replace(s, ' ')
        telefones = extrair_telefones(texto_sem_servico)

    reg['telefones'].extend(t for t in telefones if t not in reg['telefones'])
    if servicos and not reg['servico']:
        reg['servico'] = ' + '.join(dict.fromkeys(servicos))
    if not reg['operadora'] and (reg['nic'] or re.search(r'\b(MEO|MXO)', texto, re.I)):
        reg['operadora'] = 'MEO'


def interpretar_misto(texto, reg, rotulo):
    """Texto livre com varios dados misturados: extrai o que reconhece e guarda o original."""
    m = RE_EMAIL.search(texto)
    if m and not reg['email']:
        reg['email'] = m.group(0).lower()
    m = RE_NIF_ROTULO.search(texto)
    if m and not reg['nif']:
        reg['nif'] = m.group(1)
    m = RE_CP.search(texto)
    if m and not reg['cp']:
        reg['cp'] = f'{m.group(1)}-{m.group(2)}'
    reg['telefones'].extend(t for t in extrair_telefones(texto) if t not in reg['telefones'])
    if not reg['nome']:
        m = re.match(r"^\s*(?:nome\s*:\s*)?([A-Za-zÀ-ÿ'.\s]+?)\s*(?:[;|,\-–/:(]|\d|$)", texto, re.I)
        if m and len(m.group(1).split()) >= 2:
            reg['nome'] = limpar_nome(m.group(1))
    reg['outros'].append(f'{rotulo}: {texto}')


def interpretar_linha(valores, papeis, rotulos, origem):
    reg = registo_vazio(origem)
    for col, valor in enumerate(valores):
        texto = celula_texto(valor)
        if not texto:
            continue
        papel = papeis.get(col, 'outros')
        rotulo = rotulos[col]

        if papel == 'meo' or (papel in ('outros', 'misto', 'nome', 'texto') and RE_MEO.search(texto)):
            interpretar_meo(texto, reg)
            continue
        if papel in PAPEIS_UNICOS and papel != 'data' and reg.get(papel):
            reg['outros'].append(f'{rotulo}: {texto}')
            continue

        if papel == 'telefone':
            tels = extrair_telefones(texto)
            reg['telefones'].extend(t for t in tels if t not in reg['telefones'])
            if not tels:
                reg['outros'].append(f'{rotulo}: {texto}')
        elif papel == 'nome':
            reg['nome'] = limpar_nome(texto)
        elif papel == 'morada':
            reg['morada'] = re.sub(r'\s+', ' ', texto)
            m = RE_CP.search(texto)
            if m and not reg['cp']:
                reg['cp'] = f'{m.group(1)}-{m.group(2)}'
        elif papel == 'localidade':
            reg['localidade'] = texto
        elif papel == 'cp':
            m = RE_CP.search(texto)
            if m:
                reg['cp'] = f'{m.group(1)}-{m.group(2)}'
            else:
                m4 = re.search(r'\b(\d{4})\b', texto)
                if m4 and len(texto.strip()) <= 10:
                    reg['cp'] = m4.group(1)
                else:
                    tels = extrair_telefones(texto)
                    if tels:
                        reg['telefones'].extend(t for t in tels if t not in reg['telefones'])
                    else:
                        reg['outros'].append(f'{rotulo}: {texto}')
        elif papel == 'nif':
            reg['nif'] = re.sub(r'\D', '', texto) or texto
        elif papel == 'nic':
            reg['nic'] = texto
        elif papel == 'email':
            m = RE_EMAIL.search(texto)
            reg['email'] = m.group(0).lower() if m else texto
        elif papel == 'data':
            convertida = converter_data(valor)
            if convertida is not None and reg['data'] is None:
                reg['data'] = convertida
            else:
                reg['outros'].append(f'{rotulo}: {texto}')
        elif papel == 'operadora':
            reg['operadora'] = texto
        elif papel == 'servico':
            reg['servico'] = texto
        elif papel == 'misto':
            interpretar_misto(texto, reg, rotulo)
        else:
            reg['outros'].append(f'{rotulo}: {texto}')
    return reg


def processar_folha(df, origem):
    df = df.dropna(how='all').dropna(axis=1, how='all')
    if df.empty:
        return [], 0
    df = df.reset_index(drop=True)
    df.columns = range(df.shape[1])

    idx_cab, cabecalho = detetar_cabecalho(df)
    if idx_cab is not None:
        df = df.iloc[idx_cab + 1:].reset_index(drop=True)
    linhas_dados = len(df)
    if df.empty:
        return [], 0

    papeis = decidir_papeis(df, cabecalho)
    rotulos = [(cabecalho[c] if cabecalho and cabecalho[c] else f'Coluna {letra_coluna(c)}')
               for c in df.columns]

    registos = []
    for valores in df.itertuples(index=False, name=None):
        reg = interpretar_linha(valores, papeis, rotulos, origem)
        if reg['nome'] or reg['telefones'] or reg['nif'] or reg['outros'] or reg['morada']:
            registos.append(reg)
    return registos, linhas_dados


# ─── Construcao da tabela final ────────────────────────────────────────────────

def _para_inteiro_se_possivel(serie):
    preenchidos = serie.dropna()
    preenchidos = preenchidos[preenchidos.astype(str) != '']
    if len(preenchidos) and preenchidos.astype(str).str.fullmatch(r'\d{1,15}').all():
        return pd.to_numeric(serie.replace('', pd.NA), errors='coerce').astype('Int64')
    return serie.replace('', None)


def construir_tabela(registos, prefixo351):
    linhas = []
    for r in registos:
        nome = r['nome']
        empresa = e_empresa(nome)
        nome_fmt = nome if empresa else nome_em_titulo(nome)
        tels = [(f'+351{tel}' if prefixo351 else tel) for tel in r['telefones']]
        linha = {
            'Nome Completo': nome_fmt,
            'Contacto': ', '.join(tels),
            'NIF': r['nif'],
            'Nº Cliente': r['nic'],
            'Operadora': r['operadora'],
            'Serviço Atual': r['servico'],
            'Data Fidelização': r['data'],
            'Morada': r['morada'],
            'Localidade': r['localidade'],
            'Código Postal': r['cp'],
            'Email': r['email'],
            'Tipo Cliente': tipo_cliente(r['nif'], nome),
            'Outros Dados': ' | '.join(r['outros']),
            'Origem': r['origem'],
            '_n_tel': len(r['telefones']),
        }
        linhas.append(linha)

    df = pd.DataFrame(linhas)
    if df.empty:
        return df

    for c in ('NIF', 'Nº Cliente'):
        if c in df.columns:
            df[c] = _para_inteiro_se_possivel(df[c])
    if 'Data Fidelização' in df.columns:
        df['Data Fidelização'] = pd.to_datetime(df['Data Fidelização'], errors='coerce')

    # Remover colunas opcionais totalmente vazias (mantem Nome Completo e Contacto)
    obrigatorias = {'Nome Completo', 'Contacto', '_n_tel'}
    for c in list(df.columns):
        if c in obrigatorias:
            continue
        serie = df[c]
        if serie.isna().all() or (serie.dtype == object and serie.fillna('').astype(str).str.strip().eq('').all()):
            df = df.drop(columns=c)
    if 'Origem' in df.columns and df['Origem'].nunique() <= 1:
        df = df.drop(columns='Origem')
    return df


def remover_duplicados(df):
    """Mesmo Nome + mesmo primeiro Contacto = duplicado: fica a linha com mais informacao."""
    if df.empty:
        return df, 0
    nome = df['Nome Completo'].fillna('').astype(str).map(lambda s: re.sub(r'\s+', '', sem_acentos(s).upper()))
    col_tel = 'Contacto' if 'Contacto' in df.columns else ''
    tel = df[col_tel].astype(str).replace({'<NA>': '', 'nan': '', 'None': ''}) if col_tel else ''
    primeiro_tel = tel.map(lambda t: t.split(',')[0].strip() if t else '')
    nif = df['NIF'].astype(str).replace({'<NA>': '', 'nan': '', 'None': ''}) if 'NIF' in df.columns else ''

    chave = nome + '|' + primeiro_tel
    sem_tel = primeiro_tel.eq('')
    chave = chave.where(~sem_tel, nome + '|NIF:' + nif)
    unica = (nome.eq('') & sem_tel) | chave.eq('|NIF:')
    chave = chave.where(~unica, '__unica_' + df.index.astype(str))

    riqueza = df.replace('', pd.NA).notna().sum(axis=1)
    bonus_tel = (df['Contacto'].fillna('').astype(str).str.strip().ne('')).astype(int) * 10
    score = riqueza + bonus_tel
    ordem = score.sort_values(ascending=False, kind='mergesort').index
    mantidos = chave.loc[ordem].drop_duplicates(keep='first').index
    resultado = df.loc[sorted(mantidos)]
    return resultado, len(df) - len(resultado)


# ─── Exportacao ────────────────────────────────────────────────────────────────

def _escrever_folha(writer, df, nome_folha):
    df.to_excel(writer, sheet_name=nome_folha, index=False)
    livro, folha = writer.book, writer.sheets[nome_folha]
    fmt_cab = livro.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
                                'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    fmt_zebra = livro.add_format({'bg_color': '#EAF1F8'})
    for i, coluna in enumerate(df.columns):
        folha.write(0, i, coluna, fmt_cab)
        amostra = [celula_texto(v) for v in df[coluna].head(500)]
        largura = max([len(str(coluna))] + [len(v) for v in amostra]) + 2
        folha.set_column(i, i, min(max(largura, 10), 50))
    folha.set_row(0, 24)
    folha.freeze_panes(1, 0)
    if len(df):
        folha.autofilter(0, 0, len(df), len(df.columns) - 1)
        folha.conditional_format(1, 0, len(df), len(df.columns) - 1,
                                 {'type': 'formula', 'criteria': '=MOD(ROW(),2)=0', 'format': fmt_zebra})


def gravar_excel(caminho, df_com, df_sem):
    opcoes = {'options': {'strings_to_formulas': False, 'strings_to_urls': False,
                          'strings_to_numbers': False}}
    with pd.ExcelWriter(caminho, engine='xlsxwriter', date_format='dd/mm/yyyy',
                        datetime_format='dd/mm/yyyy', engine_kwargs=opcoes) as writer:
        _escrever_folha(writer, df_com, 'Clientes')
        if not df_sem.empty:
            _escrever_folha(writer, df_sem, 'Sem Contacto')


def gravar_csv(caminho, df):
    saida = df.copy()
    if 'Data Fidelização' in saida.columns:
        if pd.api.types.is_datetime64_any_dtype(saida['Data Fidelização']):
            saida['Data Fidelização'] = saida['Data Fidelização'].dt.strftime('%d/%m/%Y').fillna('')
    for c in saida.columns:
        if saida[c].dtype == object:  # evita formulas maliciosas ao abrir no Excel
            saida[c] = saida[c].map(lambda v: "'" + v if isinstance(v, str) and v[:1] in ('=', '@') else v)
    saida.to_csv(caminho, index=False, encoding='utf-8')


def nome_base_seguro(nome_ficheiro):
    base = os.path.splitext(os.path.basename(nome_ficheiro))[0]
    base = re.sub(r'[^\w\- ]', '', base).strip()[:60]
    return base or 'ficheiro'


def preparar_preview(df):
    amostra = df.head(5).copy()
    colunas = [c for c in amostra.columns if amostra[c].notna().any()
               and amostra[c].astype(str).str.strip().replace({'<NA>': '', 'NaT': ''}).ne('').any()]
    amostra = amostra[colunas] if colunas else amostra
    if 'Data Fidelização' in amostra.columns and pd.api.types.is_datetime64_any_dtype(amostra['Data Fidelização']):
        amostra['Data Fidelização'] = amostra['Data Fidelização'].dt.strftime('%d/%m/%Y')
    amostra = amostra.astype(object).where(amostra.notna(), None)
    return [{k: (str(v) if v is not None else None) for k, v in linha.items()}
            for linha in amostra.to_dict('records')]


def gerar_resultado(registos, pasta, base, prefixo351):
    """Cria o Excel + CSV de um conjunto de registos e devolve (ficheiros, stats, df_final)."""
    df = construir_tabela(registos, prefixo351)
    if df.empty:
        raise ErroUtilizador('Não foram encontrados dados de clientes no(s) ficheiro(s).')
    df, duplicados = remover_duplicados(df)
    tem_tel = df['_n_tel'] > 0
    stats = {
        'clientes': int(len(df)),
        'com_contacto': int(tem_tel.sum()),
        'sem_contacto': int((~tem_tel).sum()),
        'multi_contacto': int((df['_n_tel'] > 1).sum()),
        'duplicados': int(duplicados),
    }
    df = df.drop(columns='_n_tel')
    df_com, df_sem = df[tem_tel.values], df[~tem_tel.values]

    nome_xlsx = f'{base}_organizado.xlsx'
    nome_csv = f'{base}_CallHub.csv'
    gravar_excel(os.path.join(pasta, nome_xlsx), df_com, df_sem)
    gravar_csv(os.path.join(pasta, nome_csv), df_com)
    return [nome_xlsx, nome_csv], stats, (df_com if not df_com.empty else df)


def limpar_ficheiros_antigos():
    try:
        if not os.path.isdir(OUTPUT_DIR):
            return
        limite = time.time() - MINUTOS_VALIDADE * 60
        for entrada in os.listdir(OUTPUT_DIR):
            caminho = os.path.join(OUTPUT_DIR, entrada)
            try:
                if os.path.getmtime(caminho) < limite:
                    if os.path.isdir(caminho):
                        shutil.rmtree(caminho, ignore_errors=True)
                    else:
                        os.remove(caminho)
            except Exception:
                pass
    except Exception:
        pass


def erro_json(mensagem, estado):
    return jsonify({'error': mensagem}), estado


# ─── Rotas Flask ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return HTML_TEMPLATE


@app.route('/transformar', methods=['POST'])
def transformar():
    token = None
    pasta = None
    try:
        limpar_ficheiros_antigos()
        ficheiros = [f for f in request.files.getlist('file') if f and f.filename]
        if not ficheiros:
            return erro_json('Nenhum ficheiro selecionado.', 400)
        for f in ficheiros:
            if f.filename.rsplit('.', 1)[-1].lower() not in EXTENSOES_PERMITIDAS or '.' not in f.filename:
                return erro_json(f'Ficheiro inválido: "{f.filename}". Usa .xlsx, .xls ou .csv.', 400)

        juntar = request.form.get('merge', 'true') != 'false'
        prefixo351 = request.form.get('prefixo351') == 'true'
        token = uuid.uuid4().hex
        pasta = os.path.join(OUTPUT_DIR, token)
        os.makedirs(pasta, exist_ok=True)

        linhas_lidas = 0
        if juntar or len(ficheiros) == 1:
            todos = []
            primeiro_base = nome_base_seguro(ficheiros[0].filename)
            for f in ficheiros:
                folhas = ler_ficheiro(f)
                for nome_folha, df in folhas:
                    origem = f'{f.filename} / {nome_folha}' if len(folhas) > 1 else f.filename
                    regs, n_linhas = processar_folha(df, origem)
                    todos.extend(regs)
                    linhas_lidas += n_linhas
                    del df
                del folhas
                gc.collect()

            if not todos:
                raise ErroUtilizador('Não foram encontrados dados de clientes no(s) ficheiro(s).')

            base = primeiro_base if len(ficheiros) == 1 else f'clientes_juntos_{datetime.now():%Y%m%d_%H%M}'
            nomes, stats, df_preview = gerar_resultado(todos, pasta, base, prefixo351)
            del todos
            gc.collect()
            downloads = [
                {'label': 'Descarregar CSV para CallHub', 'tipo': 'csv', 'url': f'/descarregar/{token}/{nomes[1]}'},
                {'label': 'Descarregar Excel', 'tipo': 'xlsx', 'url': f'/descarregar/{token}/{nomes[0]}'},
            ]
        else:
            stats = {'clientes': 0, 'com_contacto': 0, 'sem_contacto': 0, 'multi_contacto': 0, 'duplicados': 0}
            todos_nomes, df_preview, usados = [], None, set()
            for f in ficheiros:
                folhas = ler_ficheiro(f)
                registos = []
                for nome_folha, df in folhas:
                    origem = f'{f.filename} / {nome_folha}' if len(folhas) > 1 else f.filename
                    regs, n_linhas = processar_folha(df, origem)
                    registos.extend(regs)
                    linhas_lidas += n_linhas
                    del df
                del folhas
                gc.collect()

                if not registos:
                    continue
                base = nome_base_seguro(f.filename)
                while base in usados:
                    base += '_2'
                usados.add(base)
                nomes, st, df_p = gerar_resultado(registos, pasta, base, prefixo351)
                todos_nomes.extend(nomes)
                for k in stats:
                    stats[k] += st[k]
                df_preview = df_p if df_preview is None else df_preview
                del registos
                gc.collect()

            if not todos_nomes:
                raise ErroUtilizador('Não foram encontrados dados de clientes nos ficheiros.')
            nome_zip = f'clientes_organizados_{datetime.now():%Y%m%d_%H%M}.zip'
            with zipfile.ZipFile(os.path.join(pasta, nome_zip), 'w', zipfile.ZIP_DEFLATED) as zf:
                for nome in todos_nomes:
                    zf.write(os.path.join(pasta, nome), arcname=nome)
            downloads = [{'label': 'Descarregar ZIP (Excel + CSV de cada ficheiro)', 'tipo': 'zip',
                          'url': f'/descarregar/{token}/{nome_zip}'}]

        stats['linhas_lidas'] = linhas_lidas
        return jsonify({
            'stats': stats,
            'downloads': downloads,
            'preview': preparar_preview(df_preview),
            'validade_minutos': MINUTOS_VALIDADE,
        })

    except ErroUtilizador as exc:
        shutil.rmtree(pasta, ignore_errors=True)
        return erro_json(str(exc), 400)
    except Exception:
        shutil.rmtree(pasta, ignore_errors=True)
        app.logger.exception('Erro inesperado ao transformar ficheiros')
        return erro_json('Ocorreu um erro inesperado ao processar o ficheiro. '
                         'Confirma que o ficheiro abre corretamente no Excel e tenta novamente.', 500)


@app.route('/descarregar/<token>/<nome_ficheiro>')
def descarregar(token, nome_ficheiro):
    if not re.fullmatch(r'[a-f0-9]{32}', token):
        abort(404)
    pasta = os.path.join(OUTPUT_DIR, token)
    if not os.path.isfile(os.path.join(pasta, os.path.basename(nome_ficheiro))):
        return ('Este ficheiro já não está disponível (os ficheiros são apagados ao fim de '
                f'{MINUTOS_VALIDADE} minutos). Volta a carregar o ficheiro original.'), 404
    return send_from_directory(pasta, nome_ficheiro, as_attachment=True, download_name=nome_ficheiro)


@app.errorhandler(413)
def ficheiro_grande(_erro):
    return erro_json('O ficheiro é demasiado grande (máximo 50 MB).', 413)


@app.errorhandler(500)
def erro_interno(_erro):
    return erro_json('Ocorreu um erro interno no servidor. Tenta novamente.', 500)


# ─── Main ──────────────────────────────────────────────────────────────────────

def open_browser():
    """Abre o browser automaticamente apos o servidor iniciar."""
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == '__main__':
    import socket

    def get_local_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    local_ip = get_local_ip()
    port = 5000

    print('=' * 55)
    print('  Transformador de Fichas de Clientes')
    print('=' * 55)
    print(f'  Este computador:  http://127.0.0.1:{port}')
    print(f'  Rede local:       http://{local_ip}:{port}')
    print('=' * 55)
    print('  NAO feche esta janela enquanto usar!')
    print('=' * 55)

    threading.Timer(1.5, open_browser).start()
    app.run(host='0.0.0.0', port=port, debug=False)
