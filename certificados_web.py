# -*- coding: utf-8 -*-
"""
App web para montar, gerar e enviar certificados em lote.

Execute:
    source .venv_certificados/bin/activate
    python certificados_web.py
"""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import smtplib
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path

from flask import Flask, Response, jsonify, render_template_string, request, send_from_directory, url_for
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as pdf_canvas
from werkzeug.utils import secure_filename

from certificados_gui import (
    COLUMN_ALIASES,
    DEFAULT_EMAIL_BODY,
    DEFAULT_EMAIL_SUBJECT,
    IMAGE_EXTENSIONS,
    PDF_EXTENSIONS,
    SPREADSHEET_EXTENSIONS,
    build_email_message,
    connect_smtp,
    email_is_valid,
    find_column,
    load_table,
    parse_hex_color,
    prepare_participants,
    replace_placeholders,
    safe_filename,
    table_headers,
    ProcessSettings,
)


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "web_certificados_data"
UPLOAD_DIR = DATA_DIR / "uploads"
PREVIEW_DIR = DATA_DIR / "previews"
OUTPUT_DIR = DATA_DIR / "outputs"
DEFAULT_TEMPLATE = BASE_DIR / "Certificado_Gabrielly_Mayara.pdf"

PDF_DPI = 120

FONT_MAP = {
    "Arial": {
        (False, False): "Helvetica",
        (True, False): "Helvetica-Bold",
        (False, True): "Helvetica-Oblique",
        (True, True): "Helvetica-BoldOblique",
    },
    "Times New Roman": {
        (False, False): "Times-Roman",
        (True, False): "Times-Bold",
        (False, True): "Times-Italic",
        (True, True): "Times-BoldItalic",
    },
    "Helvetica": {
        (False, False): "Helvetica",
        (True, False): "Helvetica-Bold",
        (False, True): "Helvetica-Oblique",
        (True, True): "Helvetica-BoldOblique",
    },
}

PIL_FONT_MAP = {
    "Arial": {
        (False, False): "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        (True, False): "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        (False, True): "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
        (True, True): "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
    },
    "Times New Roman": {
        (False, False): "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        (True, False): "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
        (False, True): "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf",
        (True, True): "/usr/share/fonts/truetype/dejavu/DejaVuSerif-BoldItalic.ttf",
    },
}


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024


HTML = r"""
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gerador de Certificados</title>
  <style>
    :root {
      --bg: #f6f7f9;
      --panel: #ffffff;
      --line: #d8dde6;
      --text: #17202a;
      --muted: #5d6877;
      --accent: #2563eb;
      --accent-2: #0f766e;
      --danger: #b42318;
      --shadow: 0 1px 3px rgba(20, 26, 35, .08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14px;
    }
    header {
      height: 56px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 18px;
      border-bottom: 1px solid var(--line);
      background: #fff;
    }
    h1 {
      margin: 0;
      font-size: 18px;
      letter-spacing: 0;
    }
    main {
      height: calc(100vh - 56px);
      display: grid;
      grid-template-columns: 310px minmax(480px, 1fr) 340px;
      gap: 12px;
      padding: 12px;
    }
    section {
      min-width: 0;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .section-head {
      height: 42px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 12px;
      border-bottom: 1px solid var(--line);
      font-weight: 700;
    }
    .section-body {
      padding: 12px;
      overflow: auto;
      height: calc(100% - 42px);
    }
    label {
      display: block;
      font-weight: 700;
      margin: 12px 0 6px;
    }
    input, select, textarea, button {
      font: inherit;
    }
    input[type="file"], input[type="text"], input[type="email"], input[type="password"],
    input[type="number"], input[type="color"], select, textarea {
      width: 100%;
      border: 1px solid #c8d0dc;
      border-radius: 6px;
      padding: 8px;
      background: #fff;
      color: var(--text);
    }
    input[type="color"] {
      height: 36px;
      padding: 3px;
    }
    textarea {
      resize: vertical;
      min-height: 92px;
      line-height: 1.35;
    }
    button {
      border: 1px solid #b8c2d2;
      border-radius: 6px;
      padding: 8px 10px;
      background: #fff;
      color: var(--text);
      cursor: pointer;
      font-weight: 700;
    }
    button.primary {
      background: var(--accent);
      border-color: var(--accent);
      color: #fff;
    }
    button.secondary {
      background: var(--accent-2);
      border-color: var(--accent-2);
      color: #fff;
    }
    button:disabled {
      opacity: .55;
      cursor: not-allowed;
    }
    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }
    .inline {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 10px;
    }
    .inline input[type="checkbox"] {
      width: auto;
    }
    .hint {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
      margin-top: 6px;
    }
    .status {
      color: var(--muted);
      font-size: 12px;
      margin-top: 6px;
      min-height: 18px;
    }
    .toolbar {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .canvas-shell {
      height: 100%;
      display: grid;
      grid-template-rows: auto 1fr;
    }
    .preview-tools {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      background: #fff;
    }
    .preview-area {
      overflow: auto;
      padding: 16px;
      background:
        linear-gradient(45deg, #eef1f5 25%, transparent 25%),
        linear-gradient(-45deg, #eef1f5 25%, transparent 25%),
        linear-gradient(45deg, transparent 75%, #eef1f5 75%),
        linear-gradient(-45deg, transparent 75%, #eef1f5 75%);
      background-size: 22px 22px;
      background-position: 0 0, 0 11px, 11px -11px, -11px 0px;
    }
    #stage {
      position: relative;
      width: min(100%, 1120px);
      margin: 0 auto;
      background: #fff;
      box-shadow: 0 12px 30px rgba(20, 26, 35, .16);
      user-select: none;
    }
    #modelImage {
      display: none;
      width: 100%;
      height: auto;
    }
    #emptyStage {
      height: 520px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: var(--muted);
      border: 1px dashed var(--line);
      background: rgba(255, 255, 255, .72);
      font-weight: 700;
    }
    .layer {
      position: absolute;
      min-width: 42px;
      min-height: 20px;
      padding: 2px 4px;
      border: 1px dashed transparent;
      outline: none;
      line-height: 1.22;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
      cursor: move;
    }
    .layer[data-bg="true"] {
      background: #fff;
    }
    .layer.selected {
      border-color: var(--accent);
      box-shadow: 0 0 0 2px rgba(37, 99, 235, .18);
    }
    .resize-handle {
      position: absolute;
      right: -5px;
      bottom: -5px;
      width: 10px;
      height: 10px;
      border: 2px solid var(--accent);
      background: #fff;
      border-radius: 50%;
      cursor: nwse-resize;
      display: none;
    }
    .layer.selected .resize-handle { display: block; }
    .layer-list {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin: 10px 0;
    }
    .layer-item {
      width: 100%;
      text-align: left;
      border-radius: 6px;
    }
    .layer-item.active {
      border-color: var(--accent);
      color: var(--accent);
      background: #eff6ff;
    }
    .segmented {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 6px;
    }
    .segmented button.active {
      border-color: var(--accent);
      color: var(--accent);
      background: #eff6ff;
    }
    .log {
      min-height: 120px;
      max-height: 220px;
      overflow: auto;
      white-space: pre-wrap;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px;
      background: #fbfcfe;
      color: #1f2937;
      margin-top: 10px;
      font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
      font-size: 12px;
    }
    .download {
      display: none;
      margin-top: 10px;
      text-decoration: none;
      color: #fff;
      background: var(--accent-2);
      border-radius: 6px;
      padding: 10px 12px;
      font-weight: 700;
      text-align: center;
    }
    @media (max-width: 1100px) {
      main {
        height: auto;
        grid-template-columns: 1fr;
      }
      section {
        min-height: 360px;
      }
      .canvas-shell {
        min-height: 640px;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>Gerador de Certificados</h1>
    <div class="toolbar">
      <button id="loadDefaultBtn" type="button">Carregar modelo da pasta</button>
      <button id="resetLayoutBtn" type="button">Layout recomendado</button>
    </div>
  </header>

  <main>
    <section>
      <div class="section-head">Arquivos</div>
      <div class="section-body">
        <label>Modelo do certificado</label>
        <input id="templateFile" type="file" accept=".pdf,.png,.jpg,.jpeg">
        <div id="templateStatus" class="status"></div>

        <label>Planilha</label>
        <input id="sheetFile" type="file" accept=".xlsx,.csv">
        <div id="sheetStatus" class="status"></div>

        <label>Coluna do nome</label>
        <select id="nameColumn"></select>

        <label>Coluna do CPF/matricula</label>
        <select id="identifierColumn"></select>

        <label>Coluna do e-mail</label>
        <select id="emailColumn"></select>

        <label>Coluna de status</label>
        <select id="statusColumn"></select>

        <div class="inline">
          <input id="skipBlocked" type="checkbox" checked>
          <span>Pular Pendente/Reprovado/Cancelado</span>
        </div>

        <label>Camadas de texto</label>
        <div id="layerList" class="layer-list"></div>
        <div class="row">
          <button id="addLayerBtn" type="button">Adicionar texto</button>
          <button id="removeLayerBtn" type="button">Remover texto</button>
        </div>
      </div>
    </section>

    <section class="canvas-shell">
      <div class="preview-tools">
        <div class="status" id="sampleStatus">A prévia usa os dados da primeira linha válida.</div>
        <div class="toolbar">
          <button id="zoomOutBtn" type="button">Menos zoom</button>
          <button id="zoomInBtn" type="button">Mais zoom</button>
        </div>
      </div>
      <div class="preview-area">
        <div id="stage">
          <div id="emptyStage">Selecione um modelo de certificado</div>
          <img id="modelImage" alt="">
        </div>
      </div>
    </section>

    <section>
      <div class="section-head">Texto e geração</div>
      <div class="section-body">
        <label>Texto selecionado</label>
        <textarea id="layerText"></textarea>
        <div class="hint">Use &lt;nome&gt;, &lt;matricula ou cpf&gt; e **negrito**.</div>

        <div class="row">
          <div>
            <label>Fonte</label>
            <select id="fontFamily">
              <option>Times New Roman</option>
              <option>Arial</option>
              <option>Helvetica</option>
            </select>
          </div>
          <div>
            <label>Tamanho</label>
            <input id="fontSize" type="number" min="0.8" max="8" step="0.1">
          </div>
        </div>

        <div class="row">
          <div>
            <label>Cor</label>
            <input id="textColor" type="color">
          </div>
          <div>
            <label>Alinhamento</label>
            <select id="textAlign">
              <option value="left">Esquerda</option>
              <option value="center">Centro</option>
              <option value="right">Direita</option>
            </select>
          </div>
        </div>

        <div class="inline">
          <input id="boldToggle" type="checkbox">
          <span>Negrito</span>
          <input id="italicToggle" type="checkbox">
          <span>Itálico</span>
        </div>

        <div class="inline">
          <input id="backgroundToggle" type="checkbox">
          <span>Fundo branco atrás do texto</span>
        </div>

        <label>E-mail remetente</label>
        <input id="senderEmail" type="email" placeholder="seuemail@gmail.com">

        <label>Senha de app/key</label>
        <input id="senderKey" type="password" placeholder="xxxx xxxx xxxx xxxx">

        <label>Assunto</label>
        <input id="emailSubject" type="text">

        <label>Mensagem</label>
        <textarea id="emailBody"></textarea>

        <div class="inline">
          <input id="sendEmail" type="checkbox">
          <span>Enviar por e-mail após gerar</span>
        </div>

        <div class="row" style="margin-top: 12px;">
          <button id="generateBtn" class="primary" type="button">Gerar ZIP</button>
          <button id="generateSendBtn" class="secondary" type="button">Gerar e enviar</button>
        </div>

        <a id="downloadLink" class="download" href="#">Baixar certificados</a>
        <div id="log" class="log"></div>
      </div>
    </section>
  </main>

  <script>
    const state = {
      templateId: null,
      sheetId: null,
      sample: {},
      preview: null,
      zoom: 1,
      selectedLayerId: null,
      layers: []
    };

    const $ = (id) => document.getElementById(id);

    function defaultLayers() {
      return [
        {
          id: crypto.randomUUID(),
          label: 'Frase principal',
          text: 'Certificamos que <nome> - <matricula ou cpf> participou com êxito do evento **Tubos de Produção e Revestimento: Desafios, Falhas e Soluções** realizado em 24/03/2026, contabilizando carga horária total de 2 horas.',
          x: 10.5, y: 33.3, width: 83, height: 15.5, fontFamily: 'Times New Roman', size: 2.45,
          color: '#000000', bold: false, italic: false, align: 'left', background: true
        },
        {
          id: crypto.randomUUID(),
          label: 'Nome em destaque',
          text: '<nome>',
          x: 25, y: 52.5, width: 50, height: 11, fontFamily: 'Times New Roman', size: 4.6,
          color: '#000000', bold: false, italic: true, align: 'center', background: true
        }
      ];
    }

    function log(message) {
      const box = $('log');
      box.textContent += message + '\n';
      box.scrollTop = box.scrollHeight;
    }

    function setStatus(id, text) {
      $(id).textContent = text || '';
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, (char) => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      }[char]));
    }

    function replacePlaceholders(text) {
      const sample = state.sample || {};
      const name = sample.nome || 'Nome da Pessoa';
      const identifier = sample.identificador || 'CPF ou matrícula';
      const email = sample.email || 'email@exemplo.com';
      return String(text || '')
        .replace(/<\s*nome\s*>/gi, name)
        .replace(/<\s*(matricula|matrícula|cpf|documento|matricula ou cpf|matrícula ou cpf|matricula\/cpf|matrícula\/cpf)\s*>/gi, identifier)
        .replace(/<\s*(email|e-mail)\s*>/gi, email);
    }

    function markupToHtml(text) {
      let html = escapeHtml(replacePlaceholders(text));
      html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
      html = html.replace(/\*([^*]+)\*/g, '<em>$1</em>');
      return html.replace(/\n/g, '<br>');
    }

    function selectedLayer() {
      return state.layers.find(layer => layer.id === state.selectedLayerId) || state.layers[0];
    }

    function selectLayer(id) {
      state.selectedLayerId = id;
      renderLayers();
      renderLayerList();
      syncPanelFromLayer();
    }

    function renderLayerList() {
      const list = $('layerList');
      list.innerHTML = '';
      state.layers.forEach((layer, index) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'layer-item' + (layer.id === state.selectedLayerId ? ' active' : '');
        button.textContent = `${index + 1}. ${layer.label}`;
        button.addEventListener('click', () => selectLayer(layer.id));
        list.appendChild(button);
      });
    }

    function fontPx(layer) {
      const stageWidth = $('stage').getBoundingClientRect().width || 1000;
      return Math.max(8, stageWidth * (Number(layer.size) || 2.5) / 100);
    }

    function renderLayers() {
      document.querySelectorAll('.layer').forEach(el => el.remove());
      if (!state.templateId) return;
      const stage = $('stage');

      state.layers.forEach(layer => {
        const el = document.createElement('div');
        el.className = 'layer' + (layer.id === state.selectedLayerId ? ' selected' : '');
        el.dataset.id = layer.id;
        el.dataset.bg = layer.background ? 'true' : 'false';
        el.style.left = `${layer.x}%`;
        el.style.top = `${layer.y}%`;
        el.style.width = `${layer.width}%`;
        el.style.height = `${layer.height || 6}%`;
        el.style.fontFamily = layer.fontFamily;
        el.style.fontSize = `${fontPx(layer)}px`;
        el.style.color = layer.color;
        el.style.fontWeight = layer.bold ? '700' : '400';
        el.style.fontStyle = layer.italic ? 'italic' : 'normal';
        el.style.textAlign = layer.align;
        el.innerHTML = markupToHtml(layer.text) + '<span class="resize-handle"></span>';
        el.addEventListener('pointerdown', startDrag);
        el.addEventListener('click', (event) => {
          event.stopPropagation();
          selectLayer(layer.id);
        });
        stage.appendChild(el);
      });
    }

    function syncPanelFromLayer() {
      const layer = selectedLayer();
      if (!layer) return;
      $('layerText').value = layer.text;
      $('fontFamily').value = layer.fontFamily;
      $('fontSize').value = layer.size;
      $('textColor').value = layer.color;
      $('textAlign').value = layer.align;
      $('boldToggle').checked = Boolean(layer.bold);
      $('italicToggle').checked = Boolean(layer.italic);
      $('backgroundToggle').checked = Boolean(layer.background);
    }

    function updateSelectedLayer(patch) {
      const layer = selectedLayer();
      if (!layer) return;
      Object.assign(layer, patch);
      renderLayers();
      renderLayerList();
    }

    function startDrag(event) {
      if (event.target.classList.contains('resize-handle')) {
        startResize(event);
        return;
      }
      const layer = state.layers.find(item => item.id === event.currentTarget.dataset.id);
      if (!layer) return;
      selectLayer(layer.id);
      event.preventDefault();
      const stage = $('stage').getBoundingClientRect();
      const startX = event.clientX;
      const startY = event.clientY;
      const initialX = layer.x;
      const initialY = layer.y;

      function move(moveEvent) {
        const dx = (moveEvent.clientX - startX) / stage.width * 100;
        const dy = (moveEvent.clientY - startY) / stage.height * 100;
        layer.x = Math.max(0, Math.min(98, initialX + dx));
        layer.y = Math.max(0, Math.min(98, initialY + dy));
        renderLayers();
      }

      function stop() {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', stop);
      }

      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', stop);
    }

    function startResize(event) {
      event.preventDefault();
      event.stopPropagation();
      const layerEl = event.target.closest('.layer');
      const layer = state.layers.find(item => item.id === layerEl.dataset.id);
      const stage = $('stage').getBoundingClientRect();
      const startX = event.clientX;
      const startY = event.clientY;
      const initialWidth = layer.width;
      const initialHeight = layer.height || 7;

      function move(moveEvent) {
        const dx = (moveEvent.clientX - startX) / stage.width * 100;
        const dy = (moveEvent.clientY - startY) / stage.height * 100;
        layer.width = Math.max(5, Math.min(100 - layer.x, initialWidth + dx));
        layer.height = Math.max(2, Math.min(100 - layer.y, initialHeight + dy));
        renderLayers();
      }

      function stop() {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', stop);
      }

      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', stop);
    }

    function setSelectOptions(select, values, selected, includeEmpty) {
      select.innerHTML = '';
      if (includeEmpty) {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = '(sem coluna)';
        select.appendChild(option);
      }
      values.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
      });
      select.value = selected || '';
    }

    async function uploadFile(endpoint, file, statusId) {
      const form = new FormData();
      form.append('file', file);
      setStatus(statusId, 'Carregando...');
      const response = await fetch(endpoint, { method: 'POST', body: form });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Falha no upload.');
      return data;
    }

    async function loadTemplateFromFile(file) {
      try {
        const data = await uploadFile('/api/template', file, 'templateStatus');
        applyTemplate(data);
      } catch (error) {
        setStatus('templateStatus', error.message);
      }
    }

    function applyTemplate(data) {
      state.templateId = data.template_id;
      state.preview = data;
      $('modelImage').src = data.preview_url;
      $('modelImage').style.display = 'block';
      $('emptyStage').style.display = 'none';
      $('modelImage').onload = () => {
        resizeStage();
        renderLayers();
      };
      setStatus('templateStatus', `${data.filename} | ${data.preview_width} x ${data.preview_height}px`);
    }

    async function loadSheetFromFile(file) {
      try {
        const data = await uploadFile('/api/spreadsheet', file, 'sheetStatus');
        state.sheetId = data.sheet_id;
        state.sample = data.sample || {};
        setSelectOptions($('nameColumn'), data.headers, data.suggestions.name, false);
        setSelectOptions($('identifierColumn'), data.headers, data.suggestions.identifier, false);
        setSelectOptions($('emailColumn'), data.headers, data.suggestions.email, false);
        setSelectOptions($('statusColumn'), data.headers, data.suggestions.status, true);
        setStatus('sheetStatus', `${data.rows} linhas | ${data.headers.length} colunas`);
        $('sampleStatus').textContent = state.sample.nome ? `Prévia: ${state.sample.nome}` : 'A prévia usa os dados da primeira linha válida.';
        renderLayers();
      } catch (error) {
        setStatus('sheetStatus', error.message);
      }
    }

    function resizeStage() {
      const image = $('modelImage');
      if (!image.naturalWidth || !image.naturalHeight) return;
      const area = document.querySelector('.preview-area');
      const available = Math.max(360, (area?.clientWidth || 900) - 32);
      const base = Math.min(1120, image.naturalWidth, available);
      $('stage').style.width = `${base * state.zoom}px`;
    }

    function payload(sendEmail) {
      return {
        template_id: state.templateId,
        sheet_id: state.sheetId,
        columns: {
          name: $('nameColumn').value,
          identifier: $('identifierColumn').value,
          email: $('emailColumn').value,
          status: $('statusColumn').value
        },
        skip_blocked_status: $('skipBlocked').checked,
        layers: state.layers,
        send_email: sendEmail,
        email: {
          sender: $('senderEmail').value,
          key: $('senderKey').value,
          subject: $('emailSubject').value,
          body: $('emailBody').value
        }
      };
    }

    async function generate(sendEmail) {
      $('downloadLink').style.display = 'none';
      $('log').textContent = '';
      if (!state.templateId || !state.sheetId) {
        log('Selecione o modelo e a planilha antes de gerar.');
        return;
      }
      if (sendEmail && !confirm('Enviar e-mails agora?')) return;

      $('generateBtn').disabled = true;
      $('generateSendBtn').disabled = true;
      log('Gerando certificados...');

      try {
        const response = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload(sendEmail))
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Falha ao gerar certificados.');

        (data.logs || []).forEach(log);
        log(`Finalizado. Gerados: ${data.generated} | Enviados: ${data.sent} | Falhas: ${data.failed}`);
        $('downloadLink').href = data.download_url;
        $('downloadLink').style.display = 'block';
      } catch (error) {
        log(`Erro: ${error.message}`);
      } finally {
        $('generateBtn').disabled = false;
        $('generateSendBtn').disabled = false;
      }
    }

    function wireControls() {
      $('templateFile').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) loadTemplateFromFile(file);
      });
      $('sheetFile').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) loadSheetFromFile(file);
      });
      $('resetLayoutBtn').addEventListener('click', () => {
        state.layers = defaultLayers();
        state.selectedLayerId = state.layers[0].id;
        renderLayerList();
        renderLayers();
        syncPanelFromLayer();
      });
      $('loadDefaultBtn').addEventListener('click', async () => {
        try {
          const response = await fetch('/api/default-template', { method: 'POST' });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error);
          applyTemplate(data);
        } catch (error) {
          setStatus('templateStatus', error.message);
        }
      });
      $('addLayerBtn').addEventListener('click', () => {
        const layer = {
          id: crypto.randomUUID(),
          label: `Texto ${state.layers.length + 1}`,
          text: '<nome>',
          x: 35, y: 45, width: 30, height: 7, fontFamily: 'Arial', size: 2.5,
          color: '#000000', bold: false, italic: false, align: 'center', background: false
        };
        state.layers.push(layer);
        selectLayer(layer.id);
      });
      $('removeLayerBtn').addEventListener('click', () => {
        if (state.layers.length <= 1) return;
        state.layers = state.layers.filter(layer => layer.id !== state.selectedLayerId);
        state.selectedLayerId = state.layers[0].id;
        renderLayerList();
        renderLayers();
        syncPanelFromLayer();
      });
      $('layerText').addEventListener('input', () => updateSelectedLayer({ text: $('layerText').value }));
      $('fontFamily').addEventListener('change', () => updateSelectedLayer({ fontFamily: $('fontFamily').value }));
      $('fontSize').addEventListener('input', () => updateSelectedLayer({ size: Number($('fontSize').value) || 2.5 }));
      $('textColor').addEventListener('input', () => updateSelectedLayer({ color: $('textColor').value }));
      $('textAlign').addEventListener('change', () => updateSelectedLayer({ align: $('textAlign').value }));
      $('boldToggle').addEventListener('change', () => updateSelectedLayer({ bold: $('boldToggle').checked }));
      $('italicToggle').addEventListener('change', () => updateSelectedLayer({ italic: $('italicToggle').checked }));
      $('backgroundToggle').addEventListener('change', () => updateSelectedLayer({ background: $('backgroundToggle').checked }));
      $('zoomInBtn').addEventListener('click', () => {
        state.zoom = Math.min(1.8, state.zoom + 0.1);
        resizeStage();
        renderLayers();
      });
      $('zoomOutBtn').addEventListener('click', () => {
        state.zoom = Math.max(0.55, state.zoom - 0.1);
        resizeStage();
        renderLayers();
      });
      $('generateBtn').addEventListener('click', () => generate(false));
      $('generateSendBtn').addEventListener('click', () => {
        $('sendEmail').checked = true;
        generate(true);
      });
      window.addEventListener('resize', renderLayers);
    }

    async function tryLoadDefaultTemplate() {
      try {
        const response = await fetch('/api/default-template', { method: 'POST' });
        if (!response.ok) return;
        applyTemplate(await response.json());
      } catch (_error) {
        return;
      }
    }

    state.layers = defaultLayers();
    state.selectedLayerId = state.layers[0].id;
    $('emailSubject').value = 'Seu certificado';
    $('emailBody').value = 'Ola, <nome>!\\n\\nSegue em anexo o seu certificado.\\n\\nAtenciosamente,\\nEquipe organizadora.';
    wireControls();
    renderLayerList();
    syncPanelFromLayer();
    tryLoadDefaultTemplate();
  </script>
</body>
</html>
"""


def ensure_dirs() -> None:
    for folder in (UPLOAD_DIR, PREVIEW_DIR, OUTPUT_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def json_error(message: str, status: int = 400):
    return jsonify({"error": message}), status


def check_basic_auth() -> bool:
    expected_password = os.getenv("CERT_WEB_PASSWORD", "").strip()
    if not expected_password:
        return True

    auth = request.authorization
    if not auth:
        return False

    expected_user = os.getenv("CERT_WEB_USER", "gepetro")
    return auth.username == expected_user and auth.password == expected_password


def auth_required():
    return Response(
        "Acesso restrito ao GEPetro.",
        401,
        {"WWW-Authenticate": 'Basic realm="GEPetro Certificados"'},
    )


@app.before_request
def protect_public_app():
    if request.path == "/health":
        return None
    if not check_basic_auth():
        return auth_required()
    return None


@app.get("/health")
def health():
    return jsonify({"ok": True})


def allowed_suffix(path: str, suffixes: set[str]) -> bool:
    return Path(path).suffix.lower() in suffixes


def save_file_upload(upload, suffixes: set[str], prefix: str) -> Path:
    if upload is None or not upload.filename:
        raise ValueError("Nenhum arquivo foi enviado.")
    if not allowed_suffix(upload.filename, suffixes):
        raise ValueError(f"Formato invalido: {Path(upload.filename).suffix}")

    ensure_dirs()
    filename = secure_filename(upload.filename) or f"{prefix}{Path(upload.filename).suffix}"
    destination = UPLOAD_DIR / f"{prefix}_{uuid.uuid4().hex}_{filename}"
    upload.save(destination)
    return destination


def metadata_path(template_id: str) -> Path:
    return UPLOAD_DIR / f"{template_id}.json"


def write_template_metadata(metadata: dict) -> None:
    metadata_path(metadata["template_id"]).write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def read_template_metadata(template_id: str) -> dict:
    path = metadata_path(template_id)
    if not path.exists():
        raise FileNotFoundError("Modelo nao encontrado.")
    return json.loads(path.read_text(encoding="utf-8"))


def render_pdf_preview(pdf_path: Path, preview_path: Path) -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        prefix = Path(temp_dir) / "preview"
        command = [
            "pdftoppm",
            "-png",
            "-singlefile",
            "-f",
            "1",
            "-l",
            "1",
            "-r",
            str(PDF_DPI),
            str(pdf_path),
            str(prefix),
        ]
        result = subprocess.run(command, check=False, capture_output=True, text=True)
        if result.returncode != 0:
            detail = result.stderr.strip() or "pdftoppm falhou"
            raise RuntimeError(f"Nao foi possivel gerar a previa do PDF: {detail}")
        shutil.copyfile(prefix.with_suffix(".png"), preview_path)


def create_template_metadata(source_path: Path, original_name: str) -> dict:
    ensure_dirs()
    template_id = uuid.uuid4().hex
    suffix = source_path.suffix.lower()
    preview_filename = f"{template_id}.png"
    preview_path = PREVIEW_DIR / preview_filename

    if suffix in PDF_EXTENSIONS:
        reader = PdfReader(str(source_path))
        if not reader.pages:
            raise ValueError("O PDF nao possui paginas.")
        page = reader.pages[0]
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        render_pdf_preview(source_path, preview_path)
        kind = "pdf"
    elif suffix in IMAGE_EXTENSIONS:
        with Image.open(source_path) as image:
            converted = image.convert("RGB")
            converted.save(preview_path, "PNG")
            page_width, page_height = converted.size
        kind = "image"
    else:
        raise ValueError("Modelo deve ser PDF, PNG ou JPG.")

    with Image.open(preview_path) as preview:
        preview_width, preview_height = preview.size

    metadata = {
        "template_id": template_id,
        "filename": original_name,
        "path": str(source_path),
        "kind": kind,
        "page_width": page_width,
        "page_height": page_height,
        "preview_filename": preview_filename,
        "preview_width": preview_width,
        "preview_height": preview_height,
    }
    write_template_metadata(metadata)
    return metadata


def template_response(metadata: dict) -> dict:
    return {
        "template_id": metadata["template_id"],
        "filename": metadata["filename"],
        "kind": metadata["kind"],
        "page_width": metadata["page_width"],
        "page_height": metadata["page_height"],
        "preview_width": metadata["preview_width"],
        "preview_height": metadata["preview_height"],
        "preview_url": url_for("serve_preview", filename=metadata["preview_filename"]),
    }


def save_spreadsheet(upload) -> dict:
    path = save_file_upload(upload, SPREADSHEET_EXTENSIONS, "planilha")
    records = load_table(path)
    if not records:
        raise ValueError("A planilha esta vazia.")

    headers = table_headers(records)
    suggestions = {
        "name": find_column(headers, COLUMN_ALIASES["name"]),
        "identifier": find_column(headers, COLUMN_ALIASES["identifier"]),
        "email": find_column(headers, COLUMN_ALIASES["email"]),
        "status": find_column(headers, COLUMN_ALIASES["status"]),
    }

    participants, _ = prepare_participants(
        records,
        suggestions["name"],
        suggestions["email"],
        suggestions["identifier"],
        suggestions["status"],
        True,
    )
    sample = {}
    if participants:
        first = participants[0]
        sample = {"nome": first.name, "identificador": first.identifier, "email": first.email}

    sheet_id = uuid.uuid4().hex
    sheet_metadata = {
        "sheet_id": sheet_id,
        "filename": Path(path).name,
        "path": str(path),
    }
    (UPLOAD_DIR / f"{sheet_id}.sheet.json").write_text(
        json.dumps(sheet_metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return {
        "sheet_id": sheet_id,
        "rows": len(records),
        "headers": headers,
        "suggestions": suggestions,
        "sample": sample,
    }


def read_sheet_path(sheet_id: str) -> Path:
    path = UPLOAD_DIR / f"{sheet_id}.sheet.json"
    if not path.exists():
        raise FileNotFoundError("Planilha nao encontrada.")
    metadata = json.loads(path.read_text(encoding="utf-8"))
    return Path(metadata["path"])


def reportlab_font(family: str, bold: bool, italic: bool) -> str:
    family_map = FONT_MAP.get(family) or FONT_MAP["Arial"]
    return family_map[(bool(bold), bool(italic))]


def pil_font_path(family: str, bold: bool, italic: bool) -> str:
    family_map = PIL_FONT_MAP.get(family) or PIL_FONT_MAP["Arial"]
    path = family_map[(bool(bold), bool(italic))]
    if os.path.exists(path):
        return path
    return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def parse_inline_segments(text: str, base_bold: bool, base_italic: bool) -> list[dict]:
    segments = []
    bold = bool(base_bold)
    italic = bool(base_italic)
    buffer = []
    index = 0

    def flush():
        if buffer:
            segments.append({"text": "".join(buffer), "bold": bold, "italic": italic})
            buffer.clear()

    while index < len(text):
        if text.startswith("**", index):
            flush()
            bold = not bold
            index += 2
            continue
        if text[index] == "*":
            flush()
            italic = not italic
            index += 1
            continue
        buffer.append(text[index])
        index += 1
    flush()
    return [segment for segment in segments if segment["text"]]


def split_segments_by_paragraph(segments: list[dict]) -> list[list[dict]]:
    paragraphs = [[]]
    for segment in segments:
        parts = segment["text"].split("\n")
        for index, part in enumerate(parts):
            if index:
                paragraphs.append([])
            if part:
                paragraphs[-1].append({**segment, "text": part})
    return paragraphs


def segment_tokens(segments: list[dict]) -> list[dict]:
    tokens = []
    for segment in segments:
        for token in re.findall(r"\S+|\s+", segment["text"]):
            tokens.append({**segment, "text": token})
    return tokens


def trim_tokens(tokens: list[dict]) -> list[dict]:
    trimmed = list(tokens)
    while trimmed and trimmed[0]["text"].isspace():
        trimmed.pop(0)
    while trimmed and trimmed[-1]["text"].isspace():
        trimmed.pop()
    return trimmed


def measure_pdf_token(token: dict, family: str, font_size: float) -> float:
    from reportlab.pdfbase import pdfmetrics

    font_name = reportlab_font(family, token["bold"], token["italic"])
    return pdfmetrics.stringWidth(token["text"], font_name, font_size)


def wrap_pdf_segments(segments: list[dict], family: str, font_size: float, max_width: float) -> list[list[dict]]:
    lines = []
    for paragraph in split_segments_by_paragraph(segments):
        current = []
        current_width = 0.0
        tokens = segment_tokens(paragraph)
        if not tokens:
            lines.append([])
            continue

        for token in tokens:
            width = measure_pdf_token(token, family, font_size)
            if token["text"].isspace() and not current:
                continue
            if current and current_width + width > max_width and not token["text"].isspace():
                lines.append(trim_tokens(current))
                current = []
                current_width = 0.0
            if token["text"].isspace() and not current:
                continue
            current.append(token)
            current_width += width

        lines.append(trim_tokens(current))
    return lines


def line_width_pdf(line: list[dict], family: str, font_size: float) -> float:
    return sum(measure_pdf_token(token, family, font_size) for token in line)


def draw_layers_on_pdf(template_metadata: dict, layers: list[dict], participant, output_path: Path) -> None:
    reader = PdfReader(template_metadata["path"])
    page = reader.pages[0]
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)

    packet = io.BytesIO()
    canvas = pdf_canvas.Canvas(packet, pagesize=(page_width, page_height))

    for layer in layers:
        text = replace_placeholders(str(layer.get("text", "")), participant)
        if not text.strip():
            continue

        family = layer.get("fontFamily") or "Arial"
        size_percent = float(layer.get("size") or 2.5)
        font_size = max(4.0, page_width * size_percent / 100)
        x = page_width * float(layer.get("x", 0)) / 100
        y_top = page_height - (page_height * float(layer.get("y", 0)) / 100)
        max_width = page_width * float(layer.get("width", 50)) / 100
        align = layer.get("align") or "left"
        line_height = font_size * 1.22
        color = parse_hex_color(layer.get("color") or "#000000")

        segments = parse_inline_segments(text, bool(layer.get("bold")), bool(layer.get("italic")))
        lines = wrap_pdf_segments(segments, family, font_size, max_width)
        total_height = max(1, len(lines)) * line_height

        if layer.get("background"):
            requested_height = float(layer.get("height") or 0)
            background_height = (
                page_height * requested_height / 100 if requested_height > 0 else total_height
            )
            pad = font_size * 0.22
            canvas.setFillColorRGB(1, 1, 1)
            canvas.rect(
                x - pad,
                y_top - background_height - pad * 0.25,
                max_width + pad * 2,
                background_height + pad * 1.35,
                fill=1,
                stroke=0,
            )

        canvas.setFillColorRGB(color[0] / 255, color[1] / 255, color[2] / 255)
        baseline = y_top - font_size * 0.86
        for line in lines:
            width = line_width_pdf(line, family, font_size)
            line_x = x
            if align == "center":
                line_x = x + max(0, (max_width - width) / 2)
            elif align == "right":
                line_x = x + max(0, max_width - width)

            cursor = line_x
            for token in line:
                font_name = reportlab_font(family, token["bold"], token["italic"])
                canvas.setFont(font_name, font_size)
                canvas.drawString(cursor, baseline, token["text"])
                cursor += measure_pdf_token(token, family, font_size)
            baseline -= line_height

    canvas.save()
    packet.seek(0)

    overlay = PdfReader(packet).pages[0]
    page.merge_page(overlay)
    writer = PdfWriter()
    writer.add_page(page)
    for extra_page in reader.pages[1:]:
        writer.add_page(extra_page)
    with output_path.open("wb") as file:
        writer.write(file)


def measure_pil_token(draw, token: dict, family: str, font_size: int) -> int:
    font = ImageFont.truetype(pil_font_path(family, token["bold"], token["italic"]), font_size)
    bbox = draw.textbbox((0, 0), token["text"], font=font)
    return int(bbox[2] - bbox[0])


def wrap_pil_segments(draw, segments: list[dict], family: str, font_size: int, max_width: int) -> list[list[dict]]:
    lines = []
    for paragraph in split_segments_by_paragraph(segments):
        current = []
        current_width = 0
        tokens = segment_tokens(paragraph)
        if not tokens:
            lines.append([])
            continue
        for token in tokens:
            width = measure_pil_token(draw, token, family, font_size)
            if token["text"].isspace() and not current:
                continue
            if current and current_width + width > max_width and not token["text"].isspace():
                lines.append(trim_tokens(current))
                current = []
                current_width = 0
            if token["text"].isspace() and not current:
                continue
            current.append(token)
            current_width += width
        lines.append(trim_tokens(current))
    return lines


def line_width_pil(draw, line: list[dict], family: str, font_size: int) -> int:
    return sum(measure_pil_token(draw, token, family, font_size) for token in line)


def draw_layers_on_image(template_metadata: dict, layers: list[dict], participant, output_path: Path) -> None:
    with Image.open(template_metadata["path"]) as original:
        image = original.convert("RGB")

    draw = ImageDraw.Draw(image)
    image_width, image_height = image.size

    for layer in layers:
        text = replace_placeholders(str(layer.get("text", "")), participant)
        if not text.strip():
            continue
        family = layer.get("fontFamily") or "Arial"
        size_percent = float(layer.get("size") or 2.5)
        font_size = max(8, int(image_width * size_percent / 100))
        x = int(image_width * float(layer.get("x", 0)) / 100)
        y = int(image_height * float(layer.get("y", 0)) / 100)
        max_width = max(1, int(image_width * float(layer.get("width", 50)) / 100))
        align = layer.get("align") or "left"
        color = parse_hex_color(layer.get("color") or "#000000")
        line_height = int(font_size * 1.22)

        segments = parse_inline_segments(text, bool(layer.get("bold")), bool(layer.get("italic")))
        lines = wrap_pil_segments(draw, segments, family, font_size, max_width)
        total_height = max(1, len(lines)) * line_height

        if layer.get("background"):
            requested_height = float(layer.get("height") or 0)
            background_height = (
                int(image_height * requested_height / 100) if requested_height > 0 else total_height
            )
            pad = int(font_size * 0.22)
            draw.rectangle(
                [x - pad, y - pad, x + max_width + pad, y + background_height + pad],
                fill=(255, 255, 255),
            )

        cursor_y = y
        for line in lines:
            width = line_width_pil(draw, line, family, font_size)
            line_x = x
            if align == "center":
                line_x = x + max(0, (max_width - width) // 2)
            elif align == "right":
                line_x = x + max(0, max_width - width)

            cursor_x = line_x
            for token in line:
                font = ImageFont.truetype(
                    pil_font_path(family, token["bold"], token["italic"]),
                    font_size,
                )
                draw.text((cursor_x, cursor_y), token["text"], fill=color, font=font)
                cursor_x += measure_pil_token(draw, token, family, font_size)
            cursor_y += line_height

    image.save(output_path, "PDF", resolution=100.0)


def generate_certificate_file(template_metadata: dict, layers: list[dict], participant, output_path: Path) -> None:
    if template_metadata["kind"] == "pdf":
        draw_layers_on_pdf(template_metadata, layers, participant, output_path)
    else:
        draw_layers_on_image(template_metadata, layers, participant, output_path)


def make_process_settings(sender: str, key: str, subject: str, body: str) -> ProcessSettings:
    return ProcessSettings(
        spreadsheet_path="",
        template_path="",
        output_folder="",
        name_column="",
        email_column="",
        identifier_column="",
        status_column="",
        certificate_text="",
        text_x_percent=0,
        text_y_percent=0,
        text_width_percent=0,
        font_size=12,
        text_color="#000000",
        text_align="left",
        sender_email=sender,
        sender_key=key,
        smtp_host="smtp.gmail.com",
        smtp_port=465,
        email_subject=subject,
        email_body=body,
        dry_run=False,
        skip_blocked_status=True,
    )


@app.get("/")
def index():
    return render_template_string(HTML)


@app.get("/previews/<path:filename>")
def serve_preview(filename: str):
    return send_from_directory(PREVIEW_DIR, filename)


@app.get("/downloads/<path:filename>")
def serve_download(filename: str):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.post("/api/template")
def api_template():
    try:
        upload = request.files.get("file")
        source_path = save_file_upload(upload, IMAGE_EXTENSIONS.union(PDF_EXTENSIONS), "modelo")
        metadata = create_template_metadata(source_path, upload.filename)
        return jsonify(template_response(metadata))
    except Exception as error:
        return json_error(str(error))


@app.post("/api/default-template")
def api_default_template():
    try:
        if not DEFAULT_TEMPLATE.exists():
            return json_error("Modelo da pasta nao encontrado.", 404)
        ensure_dirs()
        destination = UPLOAD_DIR / f"modelo_{uuid.uuid4().hex}_{DEFAULT_TEMPLATE.name}"
        shutil.copyfile(DEFAULT_TEMPLATE, destination)
        metadata = create_template_metadata(destination, DEFAULT_TEMPLATE.name)
        return jsonify(template_response(metadata))
    except Exception as error:
        return json_error(str(error))


@app.post("/api/spreadsheet")
def api_spreadsheet():
    try:
        return jsonify(save_spreadsheet(request.files.get("file")))
    except Exception as error:
        return json_error(str(error))


@app.post("/api/generate")
def api_generate():
    try:
        data = request.get_json(force=True)
        template_id = data.get("template_id")
        sheet_id = data.get("sheet_id")
        if not template_id or not sheet_id:
            return json_error("Modelo e planilha sao obrigatorios.")

        template_metadata = read_template_metadata(template_id)
        sheet_path = read_sheet_path(sheet_id)
        records = load_table(sheet_path)
        columns = data.get("columns") or {}
        layers = data.get("layers") or []
        if not layers:
            return json_error("Adicione pelo menos uma camada de texto.")

        participants, stats = prepare_participants(
            records,
            columns.get("name") or "",
            columns.get("email") or "",
            columns.get("identifier") or "",
            columns.get("status") or "",
            bool(data.get("skip_blocked_status", True)),
        )
        if not participants:
            return json_error("Nenhuma linha valida foi encontrada na planilha.")

        batch_id = uuid.uuid4().hex
        batch_folder = OUTPUT_DIR / f"lote_{batch_id}"
        batch_folder.mkdir(parents=True, exist_ok=True)

        logs = [
            f"Linhas lidas: {stats.total_rows}",
            f"Participantes validos: {stats.valid_rows}",
        ]
        if stats.skipped_pending:
            logs.append(f"Ignorados por status: {stats.skipped_pending}")
        if stats.skipped_missing_name:
            logs.append(f"Ignorados sem nome: {stats.skipped_missing_name}")
        if stats.skipped_missing_identifier:
            logs.append(f"Ignorados sem CPF/matricula: {stats.skipped_missing_identifier}")
        if stats.skipped_missing_email:
            logs.append(f"Ignorados sem e-mail: {stats.skipped_missing_email}")
        if stats.skipped_invalid_email:
            logs.append(f"Ignorados com e-mail invalido: {stats.skipped_invalid_email}")

        send_email = bool(data.get("send_email"))
        email_data = data.get("email") or {}
        sender = str(email_data.get("sender") or "").strip()
        key = str(email_data.get("key") or "")
        subject_template = str(email_data.get("subject") or DEFAULT_EMAIL_SUBJECT)
        body_template = str(email_data.get("body") or DEFAULT_EMAIL_BODY)

        smtp = None
        if send_email:
            if not email_is_valid(sender):
                return json_error("Informe um e-mail remetente valido.")
            if not key.strip():
                return json_error("Informe a senha de app/key.")
            smtp_settings = make_process_settings(sender, key, subject_template, body_template)
            smtp = connect_smtp(smtp_settings)
            logs.append("Login no e-mail realizado.")

        generated = 0
        sent = 0
        failed = 0
        pdf_paths = []

        try:
            for participant in participants:
                filename = f"Certificado_{safe_filename(participant.name)}.pdf"
                output_path = batch_folder / filename
                counter = 2
                while output_path.exists():
                    output_path = batch_folder / f"Certificado_{safe_filename(participant.name)}_{counter}.pdf"
                    counter += 1

                try:
                    generate_certificate_file(template_metadata, layers, participant, output_path)
                    generated += 1
                    pdf_paths.append(output_path)
                except Exception as error:
                    failed += 1
                    logs.append(f"Falha ao gerar {participant.name}: {error}")
                    continue

                if send_email:
                    try:
                        subject = replace_placeholders(subject_template, participant)
                        body = replace_placeholders(body_template, participant)
                        msg = build_email_message(sender, participant.email, subject, body, output_path)
                        smtp.send_message(msg)
                        sent += 1
                    except Exception as error:
                        failed += 1
                        logs.append(f"Falha ao enviar para {participant.name}: {error}")
        finally:
            if smtp is not None:
                try:
                    smtp.quit()
                except Exception:
                    pass

        zip_name = f"certificados_{batch_id}.zip"
        zip_path = OUTPUT_DIR / zip_name
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for pdf_path in pdf_paths:
                zip_file.write(pdf_path, arcname=pdf_path.name)

        return jsonify(
            {
                "generated": generated,
                "sent": sent,
                "failed": failed,
                "logs": logs,
                "download_url": url_for("serve_download", filename=zip_name),
            }
        )
    except smtplib.SMTPAuthenticationError:
        return json_error("Falha de autenticacao no e-mail. Verifique e-mail e senha de app/key.")
    except Exception as error:
        return json_error(str(error), 500)


def main() -> int:
    ensure_dirs()
    host = os.getenv("CERT_WEB_HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "5000"))
    app.run(host=host, port=port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
