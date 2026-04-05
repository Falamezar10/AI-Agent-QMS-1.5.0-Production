"""
ИИ-Агент СМК (Agentic AI Framework)
Полностью автономный агент с использованием Tool Calling, умными правками и проактивным журналом.
"""

import uuid
import glob
import os
import hashlib
import sys
os.environ["CHROMA_TELEMETRY_DISABLED"] = "1"
import base64
from cryptography.fernet import Fernet
import chromadb
from chromadb.utils import embedding_functions
from chromadb.utils.embedding_functions import OpenAIEmbeddingFunction
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from dotenv import load_dotenv
from openai import OpenAI
from datetime import datetime
import json
import customtkinter as ctk
import threading
import tkinter as tk
import keyboard
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
import re
import textwrap
from tkinter import filedialog
import shutil
import tempfile
import requests
import subprocess
import wikipedia
import queue
import webbrowser
import win32com.client
import pythoncom
import fitz  # PyMuPDF для работы с PDF
import xml.etree.ElementTree as ET

# Настраиваем внешний вид
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Загружаем переменные окружения ДО инициализации эмбеддингов
load_dotenv()
wikipedia.set_lang("ru")  # Ищем на русском

MASTER_KEY = base64.urlsafe_b64encode(b"SMK_Enterprise_Secret_Key_32byte")
fernet = Fernet(MASTER_KEY)

def get_base_path():
    """Возвращает абсолютный путь к серверной папке. Поддерживает запуск с флагом --server"""
    import sys
    import os
    
    # 1. Проверяем, передан ли флаг --server (при запуске через умный ярлык)
    if "--server" in sys.argv:
        idx = sys.argv.index("--server")
        if len(sys.argv) > idx + 1:
            server_path = sys.argv[idx + 1]
            os.makedirs(server_path, exist_ok=True)
            return server_path
            
    # 2. Стандартное поведение (если запущено напрямую без ярлыка)
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.dirname(__file__))

def get_local_path():
    """Возвращает путь к изолированной папке профиля для конкретного экземпляра Агента"""
    import hashlib
    
    # Получаем базовый путь (серверный или папку .exe)
    base = get_base_path()
    
    # Генерируем уникальный 6-значный хэш от этого пути
    path_hash = hashlib.md5(base.encode('utf-8')).hexdigest()[:6]
    
    # Создаем уникальную папку, например: SMK_Agent_a1b2c3
    local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    app_dir = os.path.join(local_app_data, f'SMK_Agent_{path_hash}')
    
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def get_db_path():
    """Теневая репликация: стягивает серверную БД на SSD пользователя для быстрой и безопасной работы"""
    import shutil
    import chromadb.api.client

    server_db = os.path.join(get_base_path(), "smk_vector_db")
    local_db = os.path.join(get_local_path(), "local_vector_db")

    needs_pull = False
    if os.path.exists(server_db):
        server_sqlite = os.path.join(server_db, "chroma.sqlite3")
        local_sqlite = os.path.join(local_db, "chroma.sqlite3")
        if not os.path.exists(local_db) or not os.path.exists(local_sqlite):
            needs_pull = True
        elif os.path.getmtime(server_sqlite) > os.path.getmtime(local_sqlite):
            needs_pull = True

    if needs_pull:
        # Принудительно освобождаем файлы БД перед перезаписью
        try: chromadb.api.client.SharedSystemClient.clear_system_cache()
        except: pass
        try:
            shutil.rmtree(local_db, ignore_errors=True)
            shutil.copytree(server_db, local_db)
        except Exception as e:
            print(f"Ошибка репликации: {e}")

    os.makedirs(local_db, exist_ok=True)
    return local_db

def get_vault_data():
    """Чтение зашифрованного Vault с fallback на переменные окружения."""
    default_vault = {
        "openrouter_key": os.getenv("OPENROUTER_API_KEY", "").strip(),
        "groq_key": "",
        "tavily_key": "",
        "admin_password": "admin"
    }
    vault_path = os.path.join(get_base_path(), "secrets.vault")
    if not os.path.exists(vault_path):
        return default_vault
    try:
        with open(vault_path, "rb") as f:
            encrypted_data = f.read()
        decrypted_data = fernet.decrypt(encrypted_data)
        data = json.loads(decrypted_data.decode("utf-8"))
        if not isinstance(data, dict):
            return default_vault
        return {
            "openrouter_key": str(data.get("openrouter_key", default_vault["openrouter_key"])).strip(),
            "groq_key": str(data.get("groq_key", "")).strip(),
            "tavily_key": str(data.get("tavily_key", "")).strip(),
            "admin_password": str(data.get("admin_password", "admin")).strip() or "admin"
        }
    except Exception:
        return default_vault

def save_vault_data(data):
    """Сохранение зашифрованного Vault."""
    try:
        payload = {
            "openrouter_key": str(data.get("openrouter_key", "")).strip(),
            "groq_key": str(data.get("groq_key", "")).strip(),
            "tavily_key": str(data.get("tavily_key", "")).strip(),
            "admin_password": str(data.get("admin_password", "admin")).strip() or "admin"
        }
        encrypted_data = fernet.encrypt(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
        with open(os.path.join(get_base_path(), "secrets.vault"), "wb") as f:
            f.write(encrypted_data)
    except Exception:
        pass

def get_llm_client():
    """Динамический клиент LLM без глобальной инициализации."""
    vault_data = get_vault_data()
    openrouter_key = vault_data.get("openrouter_key", "").strip() or os.getenv("OPENROUTER_API_KEY", "").strip()
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=openrouter_key)

def get_cloud_ef():
    """Динамическая функция эмбеддингов без глобальной инициализации."""
    settings = load_global_settings()
    emb_model = settings.get("embedding_model", "qwen/qwen3-embedding-8b")
    vault_data = get_vault_data()
    openrouter_key = vault_data.get("openrouter_key", "").strip() or os.getenv("OPENROUTER_API_KEY", "").strip()
    if openrouter_key:
        os.environ["CHROMA_OPENAI_API_KEY"] = openrouter_key
    return OpenAIEmbeddingFunction(
        api_key=openrouter_key,
        api_base="https://openrouter.ai/api/v1",
        model_name=emb_model
    )

# ==================== ФУНКЦИИ РАБОТЫ С БАЗОЙ И ФАЙЛАМИ ====================

def get_all_paragraphs(doc):
    """Собирает все абзацы документа (включая таблицы) в единый плоский список"""
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    paras.append(p)
    return paras

def read_docx_with_indices(filepath):
    """Читает docx и возвращает текст с пронумерованными абзацами"""
    if not os.path.exists(filepath):
        return None, None
    doc = Document(filepath)
    paras = get_all_paragraphs(doc)
    result = []
    for i, p in enumerate(paras):
        text = p.text.strip()
        if text:
            result.append(f"[{i}] {text}")
    return '\n'.join(result), paras

def extract_text_from_pdf(filepath):
    """Извлекает текст из PDF-документа с текстовым слоем."""
    try:
        text_content = []
        doc = fitz.open(filepath)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("text")
            if text.strip():
                text_content.append(f"--- Страница {page_num + 1} ---\n{text.strip()}")
        doc.close()
        final_text = "\n\n".join(text_content)
        if not final_text.strip():
            return "Ошибка: Не удалось извлечь текст. Возможно, это скан без текстового слоя (OCR)."
        return final_text
    except Exception as e:
        return f"Ошибка парсинга PDF: {str(e)}"

def extract_smart_vision_and_pdf(filepath):
    """Умный Vision-роутер v1.2 для PDF и изображений с mtime-кэшем."""
    try:
        filename = os.path.basename(filepath)
        ext = filepath.lower()

        cache_dir = os.path.join(get_base_path(), ".cache")
        os.makedirs(cache_dir, exist_ok=True)
        name_without_ext = os.path.splitext(filename)[0]
        rel_path = os.path.relpath(filepath, get_base_path())
        path_hash = hashlib.md5(rel_path.encode('utf-8')).hexdigest()[:6]
        cache_path = os.path.join(cache_dir, f"{name_without_ext}_{path_hash}_vision.md")

        if os.path.exists(cache_path):
            try:
                if os.path.getmtime(cache_path) >= os.path.getmtime(filepath):
                    with open(cache_path, "r", encoding="utf-8") as f:
                        return f.read()
            except Exception:
                pass

        settings = load_global_settings()
        vision_model = settings.get("vision_model", "openai/gpt-4o-mini")

        # БЕРЕМ КЛЮЧ ИЗ ЗАШИФРОВАННОГО ХРАНИЛИЩА, А НЕ ИЗ .ENV
        vault_data = get_vault_data()
        openrouter_key = vault_data.get("openrouter_key", "").strip() or os.getenv("OPENROUTER_API_KEY", "")

        def call_vision_api(base64_image):
            if not openrouter_key:
                return "[Ошибка Vision API: не задан OPENROUTER_API_KEY]"
            try:
                system_prompt = (
                    "Ты системный аналитик и продвинутый OCR. Перед тобой страница документа, "
                    "презентации или схемы. Твоя задача:\n"
                    "1. Извлечь весь читаемый текст.\n"
                    "2. Если это блок-схема — опиши логику связей словами (что откуда куда идет).\n"
                    "3. Если таблица — выведи ее в формате Markdown.\n"
                    "Выводи только полезный текст, без лишних вступлений."
                )
                response = get_llm_client().chat.completions.create(
                    model=vision_model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {
                            "role": "user",
                            "content": [
                                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                            ]
                        }
                    ]
                )
                return response.choices[0].message.content or ""
            except Exception as e:
                return f"[Ошибка Vision API: {str(e)}]"

        final_text_blocks = []
        force_vision = "vis_index" in filename.lower()

        if ext.endswith((".png", ".jpg", ".jpeg")):
            with open(filepath, "rb") as img_file:
                b64_str = base64.b64encode(img_file.read()).decode("utf-8")
            vision_text = call_vision_api(b64_str)
            final_text_blocks.append(f"--- РАСПОЗНАНО ИЗ {filename} ---\n{vision_text}")
        elif ext.endswith(".pdf"):
            doc = fitz.open(filepath)
            try:
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    native_text = page.get_text("text").strip()

                    if force_vision:
                        route_to_vision = True
                    else:
                        text_len = len(native_text)
                        drawings = page.get_drawings()
                        has_vectors = len(drawings) > 10

                        images = page.get_image_info()
                        large_images_count = 0
                        max_img_coverage = 0.0
                        page_area = page.rect.width * page.rect.height

                        for img in images:
                            img_w = img.get("width", 0)
                            img_h = img.get("height", 0)
                            img_area = img_w * img_h
                            coverage = img_area / page_area if page_area > 0 else 0
                            if coverage > max_img_coverage:
                                max_img_coverage = coverage
                            if img_area > 40000:
                                large_images_count += 1

                        route_to_vision = False
                        if text_len < 100:
                            route_to_vision = True
                        elif max_img_coverage > 0.90:
                            route_to_vision = False
                        elif max_img_coverage > 0.25:
                            route_to_vision = True
                        elif has_vectors and large_images_count > 0:
                            route_to_vision = True

                    if route_to_vision:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        b64_str = base64.b64encode(pix.tobytes("png")).decode("utf-8")
                        vision_text = call_vision_api(b64_str)
                        final_text_blocks.append(f"--- Страница {page_num + 1} (Vision OCR) ---\n{vision_text}\n")
                    else:
                        final_text_blocks.append(f"--- Страница {page_num + 1} (Native Text) ---\n{native_text}\n")
            finally:
                doc.close()
        else:
            return "Ошибка: extract_smart_vision_and_pdf поддерживает только .pdf/.png/.jpg/.jpeg"

        full_text = "\n".join(final_text_blocks)
        with open(cache_path, "w", encoding="utf-8") as f:
            f.write(full_text)
        return full_text
    except Exception as e:
        return f"Ошибка smart vision/parsing: {str(e)}"

def extract_text_from_excel_for_rag(filepath):
    """Конвертирует Excel в плоский текст для RAG, с расклейкой объединенных ячеек."""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        all_text_lines = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # 1. Читаем значения объединенных ячеек в словарь для быстрого доступа
            merged_map = {}
            for merged_range in ws.merged_cells.ranges:
                top_left_val = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
                for r in range(merged_range.min_row, merged_range.max_row + 1):
                    for c in range(merged_range.min_col, merged_range.max_col + 1):
                        merged_map[(r, c)] = top_left_val

            headers = {}
            header_row_idx = 1

            # 2. Ищем строку заголовков
            for r in range(1, 10):
                row_vals = []
                for c_idx in range(1, ws.max_column + 1):
                    val = merged_map.get((r, c_idx), ws.cell(row=r, column=c_idx).value)
                    row_vals.append(val)
                if any(row_vals):
                    for c_idx, val in enumerate(row_vals, 1):
                        if val:
                            headers[c_idx] = str(val).replace('\n', ' ').strip()
                    header_row_idx = r
                    break

            if not headers:
                continue

            # 3. Формируем атомарные строки
            for r in range(header_row_idx + 1, ws.max_row + 1):
                row_data = []
                for c_idx, header_name in headers.items():
                    val = merged_map.get((r, c_idx), ws.cell(row=r, column=c_idx).value)
                    if val is not None and str(val).strip():
                        # Заменяем переносы внутри ячеек на пробелы, чтобы строка была монолитной
                        clean_val = str(val).replace('\n', ' ').strip()
                        row_data.append(f"{header_name}: {clean_val}")
                if row_data:
                    # Каждая строка Excel = 1 неделимый элемент
                    row_text = f"[Лист '{sheet_name}', Строка {r}] " + " | ".join(row_data)
                    all_text_lines.append(row_text)

        return "\n".join(all_text_lines)
    except Exception as e:
        return f"Ошибка парсинга Excel: {str(e)}"

def extract_text_from_graphml(filepath):
    """Парсит yEd .graphml и возвращает текстовое описание для RAG."""
    try:
        namespaces = {
            'graphml': 'http://graphml.graphdrawing.org/xmlns',
            'y': 'http://www.yworks.com/xml/graphml'
        }

        tree = ET.parse(filepath)
        root = tree.getroot()

        nodes_map = {}
        edges_list = []

        # 1) Узлы и их подписи (включая group/routing)
        for node in root.iter(f'{{{namespaces["graphml"]}}}node'):
            node_id = node.get('id')

            node_labels = []
            for data_elem in node.findall(f'./{{{namespaces["graphml"]}}}data'):
                for lbl in data_elem.findall(f'.//{{{namespaces["y"]}}}NodeLabel'):
                    text = lbl.text.strip() if lbl.text else ""
                    if text:
                        node_labels.append(text.replace('\n', ' '))

            node_label = " ".join(node_labels)
            is_group = node.find(f'./{{{namespaces["graphml"]}}}graph') is not None

            is_routing = False
            if not node_label:
                is_routing = True
                node_label = f"[точка маршрутизации {node_id}]"
            elif is_group and "[Группа]" not in node_label:
                node_label = f"[Группа] {node_label}"

            nodes_map[node_id] = {
                'label': node_label,
                'is_routing': is_routing,
                'is_group': is_group
            }

        # 2) Рёбра и тип потока
        for edge in root.iter(f'{{{namespaces["graphml"]}}}edge'):
            source_id = edge.get('source')
            target_id = edge.get('target')

            edge_label = ""
            label_elem = edge.find(f'.//{{{namespaces["y"]}}}EdgeLabel')
            if label_elem is not None and label_elem.text:
                edge_label = label_elem.text.strip().replace('\n', ' ')

            flow_type = "материальный поток"
            style_elem = edge.find(f'.//{{{namespaces["y"]}}}LineStyle')
            if style_elem is not None and style_elem.get('type') in ['dashed', 'dotted']:
                flow_type = "информационный поток"

            edges_list.append({
                'source': source_id,
                'target': target_id,
                'label': edge_label,
                'type': flow_type
            })

        # 3) Пропагация названий потоков через routing-узлы
        changed = True
        while changed:
            changed = False
            for node_id, node_data in nodes_map.items():
                if not node_data['is_routing']:
                    continue
                connected_edges = [e for e in edges_list if e['source'] == node_id or e['target'] == node_id]
                known_labels = {e['label'] for e in connected_edges if e['label'] and e['label'] != "Поток без названия"}
                if not known_labels:
                    continue
                propagated_label = " + ".join(sorted(known_labels))
                for edge in connected_edges:
                    if not edge['label'] or edge['label'] == "Поток без названия":
                        edge['label'] = propagated_label
                        changed = True

        # 4) Генерация итогового текстового описания
        lines = [f"--- ОПИСАНИЕ БИЗНЕС-ПРОЦЕССА: {os.path.basename(filepath)} ---"]

        lines.append("\n=== СПИСОК БЛОКОВ И УЗЛОВ ===")
        printed_labels = set()
        for node_data in nodes_map.values():
            if node_data['is_routing']:
                continue
            label = node_data['label']
            if label not in printed_labels:
                lines.append(f"- {label}")
                printed_labels.add(label)

        lines.append("\n=== ПОТОКИ И МАРШРУТИЗАЦИЯ ===")
        if not edges_list:
            lines.append("Связи не обнаружены.")
        else:
            for edge in edges_list:
                source_name = nodes_map.get(edge['source'], {}).get('label', f"Узел {edge['source']}")
                target_name = nodes_map.get(edge['target'], {}).get('label', f"Узел {edge['target']}")
                flow_desc = edge['label'] if edge['label'] else "Поток без названия"
                lines.append(f"[{edge['type']}] '{flow_desc}' идет ОТ '{source_name}' ---> В '{target_name}'")

        final_text = "\n".join(lines)

        # 5) Попытка сохранить markdown-копию схемы в кэш (без падения RAG при ошибке)
        try:
            cache_dir = os.path.join(get_base_path(), ".cache")
            os.makedirs(cache_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            md_path = os.path.join(cache_dir, f"{base_name}_schema.md")
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(final_text)
        except Exception:
            pass

        return final_text
    except Exception as e:
        return f"Ошибка парсинга GraphML: {str(e)}"

def extract_text_from_html_diagram(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            html = f.read()

        title_match = re.search(r'<title>(.*?)</title>', html, re.IGNORECASE)
        title = title_match.group(1).strip() if title_match else "Схема_без_названия"

        code_match = re.search(r'<textarea\s+id=["\']rawCode["\'][^>]*>(.*?)</textarea>', html, re.IGNORECASE | re.DOTALL)
        if not code_match:
            return "Ошибка: В HTML-файле не найден исходный код Mermaid (отсутствует textarea id='rawCode')."

        code = code_match.group(1).strip()
        return f"--- СХЕМА MERMAID: {title} ---\n{code}"
    except Exception as e:
        return f"Ошибка парсинга HTML-диаграммы: {str(e)}"

def transcribe_audio_logic(filename, app_instance):
    target_file = find_target_file(filename)
    if not target_file:
        return f"Ошибка: Аудиофайл '{filename}' не найден."

    def log_progress(msg):
        if app_instance is not None:
            # Убран тег "agent_msg", чтобы не было серого фона
            app_instance.after(0, lambda: app_instance.append_to_chat(f"  [Система: 🎙️ {msg}]\n"))

    temp_dir = None
    try:
        log_progress(f"Старт транскрибации: {os.path.basename(target_file)}")

        global_settings = load_global_settings()
        local_settings = load_local_settings()
        vault = get_vault_data()

        provider = global_settings.get("audio_provider", "OpenRouter")
        model = global_settings.get("audio_model", "openai/gpt-4o-audio-preview")
        chunk_mins = int(global_settings.get("audio_chunk_mins", 60))
        overlap_secs = int(global_settings.get("audio_overlap_secs", 15))

        proxies = None
        if local_settings.get("use_proxy", False):
            host = local_settings.get("proxy_host", "127.0.0.1")
            port = local_settings.get("proxy_port", "2080")
            proxies = {"http": f"socks5://{host}:{port}", "https": f"socks5://{host}:{port}"}

        # 1. Длина аудио
        probe_cmd = ["ffprobe", "-v", "quiet", "-print_format", "json", "-show_format", target_file]
        probe_result = subprocess.run(probe_cmd, capture_output=True, text=True, encoding='utf-8')
        duration_secs = float(json.loads(probe_result.stdout)['format']['duration'])
        log_progress(f"Длительность: {int(duration_secs)} сек. Нарезка на куски...")

        # 2. Нарезка
        temp_dir = os.path.join(os.path.dirname(target_file), ".temp_audio")
        os.makedirs(temp_dir, exist_ok=True)
        chunks_paths = []
        start_time = 0.0
        chunk_len_sec = chunk_mins * 60

        while start_time < duration_secs:
            out_path = os.path.join(temp_dir, f"chunk_{len(chunks_paths)}.mp3")
            ffmpeg_cmd = ["ffmpeg", "-y", "-i", target_file, "-ss", str(start_time), "-t", str(chunk_len_sec), "-c:a", "libmp3lame", "-b:a", "64k", out_path]
            subprocess.run(ffmpeg_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.exists(out_path):
                chunks_paths.append(out_path)
            start_time += (chunk_len_sec - overlap_secs)
            if chunk_len_sec <= overlap_secs:
                start_time += chunk_len_sec

        log_progress(f"Подготовлено кусков: {len(chunks_paths)}. Отправка в {provider}/{model}...")

        full_transcription = []
        # 3. Отправка
        for i, chunk_path in enumerate(chunks_paths):
            log_progress(f"Отправка куска {i + 1}/{len(chunks_paths)}")
            if provider == "Groq":
                api_key = vault.get("groq_key", "")
                if not api_key:
                    raise ValueError("Не настроен Groq API Key")
                url = "https://api.groq.com/openai/v1/audio/transcriptions"
                with open(chunk_path, "rb") as f:
                    files = {"file": (os.path.basename(chunk_path), f, "audio/mpeg")}
                    data = {"model": model, "temperature": "0.1", "response_format": "text", "language": "ru"}
                    resp = requests.post(url, headers={"Authorization": f"Bearer {api_key}"}, files=files, data=data, proxies=proxies)
                if resp.status_code == 200:
                    full_transcription.append(resp.text.strip())
                else:
                    raise ValueError(f"Ошибка Groq: {resp.text}")
            else:
                api_key = vault.get("openrouter_key", "") or os.getenv("OPENROUTER_API_KEY", "")
                with open(chunk_path, "rb") as f:
                    b64_audio = base64.b64encode(f.read()).decode('utf-8')
                prompt = "Ты профессиональный стенографист. Твоя задача - дословная расшифровка аудио.\nПРАВИЛА:\n1. Выведи ТОЛЬКО текст, который произносят люди.\n2. НИКАКИХ своих комментариев.\n3. ТРАНСКРИБИРУЙ ВЕСЬ КУСОК ДО САМОГО КОНЦА, пиши всё, что слышишь."
                payload = {
                    "model": model,
                    "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "input_audio", "input_audio": {"data": b64_audio, "format": "mp3"}}]}],
                    "temperature": 0.1,
                    "frequency_penalty": 0.5
                }
                resp = requests.post("https://openrouter.ai/api/v1/chat/completions", json=payload, headers={"Authorization": f"Bearer {api_key}"}, proxies=proxies)
                if resp.status_code == 200:
                    full_transcription.append(resp.json().get('choices', [{}])[0].get('message', {}).get('content', ''))
                else:
                    raise ValueError(f"Ошибка OpenRouter: {resp.text}")

        log_progress("Сборка финальной транскрипции...")
        final_text = "\n\n".join(full_transcription)

        # 4. Сохранение
        base_dir = os.path.dirname(target_file)
        name_without_ext = os.path.splitext(os.path.basename(target_file))[0]
        abs_path = os.path.abspath(target_file)

        # Кэш
        cache_dir = os.path.join(get_base_path(), ".cache")
        os.makedirs(cache_dir, exist_ok=True)
        path_hash = hashlib.md5(abs_path.encode("utf-8")).hexdigest()[:6]
        with open(os.path.join(cache_dir, f"{name_without_ext}_{path_hash}_audio.md"), "w", encoding="utf-8") as f:
            f.write(final_text)

        # Docx
        docx_path = os.path.join(base_dir, f"ТРАНСКРИПЦИЯ_{name_without_ext}.docx")
        doc = Document()
        doc.add_paragraph(f"Расшифровка аудио: {os.path.basename(target_file)}").style = 'Heading 1'
        doc.add_paragraph(f"Дата генерации: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n---")
        for p_text in final_text.split("\n\n"):
            if p_text.strip():
                doc.add_paragraph(p_text.strip())
        doc.save(docx_path)
        log_progress(f"Завершено. Создан документ: {os.path.basename(docx_path)}")

        threading.Thread(target=sync_vector_db, daemon=True).start()
        return f"Аудиофайл успешно расшифрован! Создан документ: [Из файла: {os.path.basename(docx_path)}]"
    except Exception as e:
        log_progress(f"Ошибка: {str(e)}")
        return f"Ошибка транскрибации: {str(e)}"
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

def convert_legacy_to_docx(input_path, output_path):
    """Конвертирует .doc/.rtf в .docx через скрытый COM-объект Word"""
    pythoncom.CoInitialize()
    word_app = None
    doc = None
    try:
        # DispatchEx для изоляции процесса
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        # Открываем только для чтения
        doc = word_app.Documents.Open(input_path, False, True, False)
        # Сохраняем как .docx
        doc.SaveAs2(output_path, FileFormat=16)
        return True, output_path
    except Exception as e:
        return False, f"Ошибка COM-конвертации: {str(e)}"
    finally:
        if doc:
            try: doc.Close(0)
            except: pass
        if word_app:
            try: word_app.Quit()
            except: pass
        pythoncom.CoUninitialize()

def find_target_file(filename):
    """Единый локатор файлов/папок с учетом настроек и black-list слов."""
    try:
        if os.path.isabs(filename) and os.path.exists(filename):
            return filename
        if os.path.exists(filename):
            return filename

        settings = load_global_settings()
        folders = settings.get("indexed_folders", ["./SMK_Docs", "./Memory"])
        excludes = [k.lower() for k in settings.get("exclude_keywords", [])]

        target_name = os.path.basename(str(filename).strip()).lower()
        if not target_name:
            return None

        def has_excluded(text):
            text_low = str(text).lower()
            return any(k and k in text_low for k in excludes)

        for folder in folders:
            if not os.path.exists(folder):
                continue

            for root, dirs, files in os.walk(folder):
                if has_excluded(root):
                    dirs[:] = []
                    continue

                dirs[:] = [d for d in dirs if not has_excluded(d)]

                for d in dirs:
                    if d.lower() == target_name:
                        return os.path.join(root, d)

                for f in files:
                    if has_excluded(f):
                        continue
                    if f.lower() == target_name:
                        return os.path.join(root, f)

        return None
    except Exception:
        return None

def read_local_file(filename):
    target_file = find_target_file(filename)
    if not target_file:
        return f"Ошибка: Файл '{filename}' не найден в разрешенных директориях."

    if os.path.isdir(target_file):
        allowed_exts = (
            '.docx', '.txt', '.md', '.pdf', '.png', '.jpg', '.jpeg',
            '.xlsx', '.xls', '.doc', '.rtf', '.graphml', '.html',
            '.mp3', '.wav', '.m4a', '.ogg'
        )
        files = [f for f in os.listdir(target_file) if f.lower().endswith(allowed_exts)]
        return f"ОШИБКА: '{filename}' - это папка. Доступные файлы внутри: {', '.join(files)}. Вызови этот инструмент заново для каждого файла по отдельности."

    # 4. Читаем сам файл
    try:
        ext = target_file.lower()
        if ext.endswith('.txt') or ext.endswith('.md'):
            with open(target_file, 'r', encoding='utf-8') as f: return f.read()
        elif ext.endswith('.docx'):
            return read_docx_with_indices(target_file)[0]
        elif ext.endswith(('.doc', '.rtf')):
            # Создаем скрытую кэш-директорию
            cache_dir = os.path.join(get_base_path(), ".cache")
            os.makedirs(cache_dir, exist_ok=True)

            # COM требует абсолютных путей
            input_abs_path = os.path.abspath(target_file)
            base_name = os.path.basename(input_abs_path)
            name_without_ext = os.path.splitext(base_name)[0]
            rel_path = os.path.relpath(input_abs_path, get_base_path())
            path_hash = hashlib.md5(rel_path.encode('utf-8')).hexdigest()[:6]
            output_abs_path = os.path.join(cache_dir, f"{name_without_ext}_{path_hash}_converted.docx")

            success, result_path = convert_legacy_to_docx(input_abs_path, output_abs_path)
            if success:
                return read_docx_with_indices(result_path)[0]
            else:
                return f"Ошибка чтения старого формата: {result_path}"
        elif ext.endswith(('.pdf', '.png', '.jpg', '.jpeg')):
            return extract_smart_vision_and_pdf(target_file)
        elif ext.endswith('.xlsx') or ext.endswith('.xls'):
            return extract_text_from_excel_for_rag(target_file)
        elif ext.endswith('.graphml'):
            return extract_text_from_graphml(target_file)
        elif ext.endswith('.html'):
            return extract_text_from_html_diagram(target_file)
        elif ext.endswith(('.mp3', '.wav', '.m4a', '.ogg')):
            # Используем АБСОЛЮТНЫЙ путь для 100% совпадения хэшей
            abs_path = os.path.abspath(target_file)
            path_hash = hashlib.md5(abs_path.encode('utf-8')).hexdigest()[:6]
            name_without_ext = os.path.splitext(os.path.basename(target_file))[0]
            
            cache_path = os.path.join(get_base_path(), ".cache", f"{name_without_ext}_{path_hash}_audio.md")

            if os.path.exists(cache_path):
                with open(cache_path, "r", encoding="utf-8") as f:
                    return f"--- РАСШИФРОВКА АУДИО ({os.path.basename(target_file)}) ---\n{f.read()}"
            
            return f"[Системная метка: Аудиофайл '{os.path.basename(target_file)}'. Текст еще НЕ расшифрован. Спроси у пользователя разрешение и вызови инструмент 'transcribe_audio_file'.]"
        else:
            return "Ошибка: Поддерживаются только форматы .txt, .md, .docx, .doc, .rtf, .pdf, .png, .jpg, .jpeg, .xlsx, .xls, .graphml, .html"
    except Exception as e: return f"Ошибка чтения файла: {str(e)}"

def chunk_text(text, chunk_size=350, overlap=50):
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end])
        start += chunk_size - overlap
    return chunks

def scan_folders_for_docs(folders):
    settings = load_global_settings()
    excludes = [k.lower() for k in settings.get("exclude_keywords", [])]
    allowed_exts = ('.docx', '.txt', '.md', '.pdf', '.png', '.jpg', '.jpeg', '.xlsx', '.xls', '.doc', '.rtf', '.graphml', '.html', '.mp3', '.wav', '.m4a', '.ogg')

    def has_excluded(text):
        text_low = str(text).lower()
        return any(k and k in text_low for k in excludes)

    # Используем SET (множество) вместо списка для автоматического удаления дубликатов!
    found_files = set() 

    for folder in folders:
        if not os.path.exists(folder):
            continue # Сканер не должен создавать папки-опечатки пользователей

        for root, dirs, files in os.walk(folder):
            root_low = root.lower()
            if '.cache' in root_low or has_excluded(root):
                dirs[:] = []
                continue

            dirs[:] = [d for d in dirs if '.cache' not in d.lower() and not has_excluded(d)]

            for filename in files:
                if filename.startswith('~$'):
                    continue
                if has_excluded(filename):
                    continue
                ext = filename.lower()
                if ext.endswith(allowed_exts):
                    # Приводим путь к абсолютному стандарту ОС, чтобы исключить дубли из-за разных слешей
                    full_path = os.path.abspath(os.path.normpath(os.path.join(root, filename)))
                    found_files.add(full_path)
                    
    # Возвращаем обратно список, как того ожидает остальной код
    return list(found_files)

def get_file_states():
    file_states_path = os.path.join(get_base_path(), "file_states.json")
    if os.path.exists(file_states_path):
        try:
            with open(file_states_path, 'r', encoding='utf-8') as f: return json.load(f)
        except: pass
    return {}

def save_file_states(states):
    try:
        file_states_path = os.path.join(get_base_path(), "file_states.json")
        with open(file_states_path, 'w', encoding='utf-8') as f: json.dump(states, f, ensure_ascii=False, indent=2)
    except: pass

def list_available_files(category="all", search_keyword=""):
    """Инструмент: Умный поиск и группировка проиндексированных файлов из file_states.json"""
    try:
        states = get_file_states()
        if not states:
            return "База файлов пуста. Подскажи пользователю нажать 'Синхронизировать базу'."
            
        ext_map = {
            "audio": ('.mp3', '.wav', '.m4a', '.ogg'),
            "excel": ('.xlsx', '.xls'),
            "word": ('.docx', '.doc', '.rtf'),
            "pdf": ('.pdf',),
            "image": ('.png', '.jpg', '.jpeg'),
            "text": ('.txt', '.md'),
            "diagram": ('.graphml', '.html')
        }

        labels = {
            "audio": "🎙️ Аудиофайлы",
            "excel": "📊 Таблицы Excel",
            "word": "📄 Word Документы",
            "pdf": "📕 PDF Документы",
            "image": "🖼️ Изображения",
            "text": "📝 Текстовые файлы",
            "diagram": "📈 Схемы и Диаграммы"
        }
        
        grouped_files = {k: [] for k in ext_map.keys()}
        grouped_files["other"] = []
        labels["other"] = "📁 Другие файлы"

        keyword = str(search_keyword).lower().strip()
        total_found = 0
        
        for path in states.keys():
            ext = os.path.splitext(path)[1].lower()
            name = os.path.basename(path)
            
            # Фильтр по ключевому слову в названии
            if keyword and keyword not in name.lower():
                continue
                
            # Определяем категорию
            matched_cat = "other"
            for cat, exts in ext_map.items():
                if ext in exts:
                    matched_cat = cat
                    break
                    
            # Фильтр по категории (если запрошена конкретная)
            if category != "all" and matched_cat != category:
                continue
                
            grouped_files[matched_cat].append(name)
            total_found += 1
            
        if total_found == 0:
            msg = "В базе не найдено файлов."
            if category != "all": msg += f" Категория: '{category}'."
            if keyword: msg += f" Искомое слово: '{keyword}'."
            return msg
            
        output_lines = [f"НАЙДЕНО ФАЙЛОВ ({total_found} шт):"]
        
        # Собираем красивый структурированный список для Агента
        for cat, files in grouped_files.items():
            if files:
                output_lines.append(f"\n{labels[cat]}:")
                unique_files = sorted(list(set(files)))
                # Ограничиваем вывод одной категории 30 файлами, чтобы не взорвать контекст
                for f in unique_files[:30]:
                    output_lines.append(f"  - {f}")
                if len(unique_files) > 30:
                    output_lines.append(f"  ... и еще {len(unique_files) - 30} файлов этой категории.")
                    
        return "\n".join(output_lines)
    except Exception as e:
        return f"Ошибка при получении списка файлов: {str(e)}"

def sync_vector_db(self=None):
    try:
        # --- ПРЕДОХРАНИТЕЛЬ: Проверяем наличие реального ключа ---
        vault_data = get_vault_data()
        raw_key = str(vault_data.get("openrouter_key", "")).strip() or os.getenv("OPENROUTER_API_KEY", "").strip()
        if not raw_key or raw_key == "sk-dummy-key":
            raise ValueError("Ожидание API-ключа. Зайдите как Админ и введите ключ в Настройках.")
        # ---------------------------------------------------------

        db_path = get_db_path()
        try:
            client = chromadb.PersistentClient(path=db_path)
            collection = client.get_or_create_collection(name="smk_docs", embedding_function=get_cloud_ef())
        except Exception as db_err:
            raise ValueError(f"Ошибка доступа к локальной БД: {db_err}. Перезапустите программу.")

        # ЭШЕЛОН ЗАЩИТЫ БАЗЫ: Гости только подключаются к БД, но не сканируют папки!
        if self is not None and getattr(self, "current_role", "guest") != "admin":
            return collection, collection.count()
        
        settings = load_global_settings()
        # пользовательские папки
        folders_to_scan = settings.get("indexed_folders", [])
        # системная папка памяти всегда должна быть
        memory_dir = os.path.join(get_base_path(), "Memory")
        os.makedirs(memory_dir, exist_ok=True)
        if memory_dir not in folders_to_scan:
            folders_to_scan.append(memory_dir)
        file_states = get_file_states()
        found_files = scan_folders_for_docs(folders_to_scan)
        
        # ВОССТАНОВЛЕННЫЕ ПЕРЕМЕННЫЕ
        new_file_states = {}
        files_to_reindex = []
        untranscribed_audio = [] # Список для оповещений

        for file_path in found_files:
            filename = os.path.basename(file_path)
            mtime = str(os.path.getmtime(file_path))
            new_file_states[file_path] = mtime
            
            # --- ЗАЩИТА ОТ ДУБЛИРОВАНИЯ И ПРОВЕРКА КЭША ---
            if file_path.lower().endswith(('.mp3', '.wav', '.m4a', '.ogg')):
                abs_path = os.path.abspath(file_path)
                path_hash = hashlib.md5(abs_path.encode('utf-8')).hexdigest()[:6]
                name_without_ext = os.path.splitext(filename)[0]
                c_path = os.path.join(get_base_path(), ".cache", f"{name_without_ext}_{path_hash}_audio.md")
                
                # Если кэша нет - добавляем в список неопознанных
                if not os.path.exists(c_path):
                    untranscribed_audio.append(filename)
                
                # КРИТИЧЕСКИ ВАЖНО: Пропускаем добавление аудио в files_to_reindex,
                # чтобы Chroma DB не засорялась текстом из кэша (для этого есть docx)
                continue
            # ----------------------------------------------

            if file_path not in file_states or file_states[file_path] != mtime:
                files_to_reindex.append((file_path, filename))
        
        current_files = set(new_file_states.keys())
        stored_files = set(file_states.keys())
        deleted_files = stored_files - current_files
        
        for file_path in deleted_files:
            try: collection.delete(where={"file_path": file_path})
            except: pass
        
        for i, (file_path, filename) in enumerate(files_to_reindex):
            if self is not None and len(files_to_reindex) > 0:
                progress = (i + 1) / len(files_to_reindex)
                current_filename = os.path.basename(file_path)
                self.after(0, lambda p=progress, f=current_filename: self.update_progress_ui(p, f))
            try:
                collection.delete(where={"file_path": file_path})
                text = read_local_file(file_path)
                if "Ошибка" in text: continue

                # ЭШЕЛОН 4: Атомарные чанки для Excel
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    # Не рубим Excel мясорубкой. Каждая строка (отбитая \n) = отдельный чанк
                    chunks = [line for line in text.split('\n') if line.strip()]
                else:
                    chunks = chunk_text(text)
                batch_docs = []
                batch_ids = []
                batch_metas = []
                
                # 1. Собираем все чанки файла в списки
                for j, chunk in enumerate(chunks):
                    if chunk.strip():
                        batch_docs.append(chunk)
                        batch_ids.append(f"{file_path}_chunk_{j}")
                        batch_metas.append({"source": filename, "file_path": file_path})
                
                # 2. Пакетная отправка (Batching) настраиваемыми пакетами
                settings = load_global_settings()
                batch_size = int(settings.get("chroma_batch_size", 100))
                for j in range(0, len(batch_docs), batch_size):
                    collection.upsert(
                        documents=batch_docs[j:j+batch_size],
                        ids=batch_ids[j:j+batch_size],
                        metadatas=batch_metas[j:j+batch_size]
                    )
            except Exception as e:
                print(f"Ошибка индексации {filename}: {e}")
                
        # Оповещение о нерасшифрованных аудио в чат
        if untranscribed_audio and self is not None:
            unique_audio = list(set(untranscribed_audio))
            display_names = ", ".join(unique_audio[:5]) + ("..." if len(unique_audio) > 5 else "")
            msg = f"\n[Система: ⚠️ В базе обнаружены нерасшифрованные аудиофайлы ({len(unique_audio)} шт.): {display_names}. Запустить транскрибацию?]\n\n"
            self.after(0, lambda m=msg: self.append_to_chat(m))

        save_file_states(new_file_states)

        if self is not None and getattr(self, "current_role", "guest") == "admin":
            import shutil
            server_db = os.path.join(get_base_path(), "smk_vector_db")
            local_db = db_path
            self.after(0, lambda: self.file_progress_label.configure(text="Отправка базы на сервер..."))
            try:
                chromadb.api.client.SharedSystemClient.clear_system_cache()
                shutil.rmtree(server_db, ignore_errors=True)
                shutil.copytree(local_db, server_db)
                client = chromadb.PersistentClient(path=local_db)
                collection = client.get_or_create_collection(name="smk_docs", embedding_function=get_cloud_ef())
            except Exception as e:
                print(f"Ошибка выгрузки БД на сервер: {e}")

        return collection, collection.count()
    finally:
        if self is not None:
            self.after(0, lambda: self.update_progress_ui(0, "Синхронизация завершена"))

# ==================== ИНСТРУМЕНТЫ АГЕНТА (ПК-РУКИ) ====================

def recall_past_conversation(query, app_instance=None):
    """Поиск по архиву текущего диалога (вытесненный контекст)"""
    if not app_instance:
        return "Ошибка: Контекст сессии не найден."
    try:
        client = chromadb.PersistentClient(path=get_db_path())
        collection = client.get_or_create_collection(name="temp_chat_memory", embedding_function=get_cloud_ef())
        results = collection.query(
            query_texts=[query],
            n_results=3,
            where={"session_id": app_instance.current_session_id}
        )
        docs = results.get('documents', [[]])[0]
        if not docs: return "В архиве старых сообщений ничего не найдено."
        return "НАЙДЕНО В АРХИВЕ:\n" + "\n---\n".join(docs)
    except Exception as e: return f"Ошибка поиска в архиве: {str(e)}"

def search_smk_knowledge_base(query):
    try:
        client = chromadb.PersistentClient(path=get_db_path())
        collection = client.get_or_create_collection(name="smk_docs", embedding_function=get_cloud_ef())
        results = collection.query(query_texts=[query], n_results=5)
        documents = results.get('documents', [[]])[0]
        sources = [meta.get('source', '') for meta in results.get('metadatas', [[]])[0]]
        
        if not documents: return "В базе знаний ничего не найдено."
        response = []
        for doc, source in zip(documents, sources):
            # Жестко задаем формат тега прямо в контексте!
            response.append(f"Источник: [Из файла: {source}]\n{doc}")
        return "\n\n---\n\n".join(response)
    except Exception as e:
        if "locked" in str(e).lower():
            return "⏳ База знаний СМК сейчас обновляется Администратором. Пожалуйста, подождите 1-2 минуты и повторите запрос."
        return f"Ошибка поиска: {str(e)}"

def web_search_tavily(query):
    """Поиск по всему интернету через Tavily"""
    api_key = get_vault_data().get("tavily_key", "").strip()
    if not api_key:
        return "Ошибка: Ключ Tavily API не настроен в Vault."

    url = "https://api.tavily.com/search"
    payload = {
        "api_key": api_key,
        "query": query,
        "search_depth": "advanced",
        "include_answer": False,
        "include_images": False,
        "max_results": 5
    }
    try:
        response = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
        response.raise_for_status()
        results = response.json().get("results", [])
        if not results:
            return "К сожалению, поиск в интернете не дал результатов."

        output = ["НАЙДЕННЫЕ МАТЕРИАЛЫ ИЗ ИНТЕРНЕТА (TAVILY):"]
        for i, res in enumerate(results, 1):
            output.append(f"--- ИСТОЧНИК {i}: {res.get('title', '')} ---")
            output.append(f"Ссылка: {res.get('url', '')}")
            output.append(f"Текст:\n{res.get('content', '')}\n")
        return "\n".join(output)
    except Exception as e:
        return f"Ошибка при поиске в интернете: {e}"

def search_wikipedia_tool(query):
    """Поиск определений и фактов в Википедии"""
    try:
        search_results = wikipedia.search(query, results=1)
        if not search_results:
            return "В Википедии ничего не найдено по этому запросу."
        page = wikipedia.page(search_results[0])
        # Берем первые 2500 символов, чтобы не перегружать контекст
        return f"--- ВИКИПЕДИЯ: {page.title} ---\n{page.summary[:2500]}\nСсылка: {page.url}"
    except wikipedia.exceptions.DisambiguationError as e:
        return f"Запрос слишком многозначный. Уточните: {e.options[:5]}"
    except Exception as e:
        return f"Ошибка поиска в Википедии: {str(e)}"

def memorize_important_fact(fact):
    try:
        memory_dir = os.path.join(get_base_path(), "Memory")
        os.makedirs(memory_dir, exist_ok=True)
        memory_file = os.path.join(memory_dir, "agent_memory.md")
        date_str = datetime.now().strftime("%d.%m.%Y %H:%M")
        if not os.path.exists(memory_file):
            with open(memory_file, "w", encoding="utf-8") as f: f.write("# Долгосрочная память ИИ-Агента\n\n")
        with open(memory_file, "a", encoding="utf-8") as f: f.write(f"- [{date_str}] {fact}\n")
        sync_vector_db()
        return f"Факт успешно сохранен и проиндексирован."
    except Exception as e: return f"Ошибка памяти: {str(e)}"

def forget_fact(query):
    try:
        memory_file = os.path.join(get_base_path(), "Memory", "agent_memory.md")
        if not os.path.exists(memory_file): return "Файл памяти пуст."
        with open(memory_file, "r", encoding="utf-8") as f: lines = f.readlines()
        prompt = f"Файл памяти:\n{''.join(lines)}\n\nУдали: '{query}'. Какую строку удалить? Верни ТОЛЬКО точный текст строки, либо 'NOT_FOUND'."
        resp = get_llm_client().chat.completions.create(model="openai/gpt-4o-mini", messages=[{"role": "user", "content": prompt}])
        line_to_delete = resp.choices[0].message.content.strip()
        if line_to_delete == "NOT_FOUND": return "Факт не найден."
        new_lines = [line for line in lines if line_to_delete not in line]
        with open(memory_file, "w", encoding="utf-8") as f: f.writelines(new_lines)
        sync_vector_db()
        return f"Факт удален."
    except Exception as e: return f"Ошибка удаления: {str(e)}"

def generate_mermaid_diagram(title: str, mermaid_code: str, app_instance=None) -> str:
    try:
        cleaned_code = (mermaid_code or "").strip()
        cleaned_code = cleaned_code.replace("```mermaid", "").replace("```", "").strip()

        safe_title = re.sub(r'[\\/*?:"<>|]', "", title or "Mermaid_Diagram").replace(" ", "_").strip("._")
        if not safe_title:
            safe_title = "Mermaid_Diagram"
        filename = f"{safe_title}.html"

        if app_instance is not None:
            output_path = app_instance.ask_save_path_sync(filename, ext=".html")
            if not output_path:
                return "Сохранение диаграммы отменено пользователем."
        else:
            output_dir = os.path.join(get_base_path(), "Созданные_Документы", "Схемы")
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, filename)

        html_content = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; color: #333; display: flex; flex-direction: column; align-items: center; padding: 20px; margin: 0; height: 100vh; box-sizing: border-box; }}
        .container {{ background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 8px 20px rgba(0,0,0,0.05); width: 100%; max-width: 95vw; display: flex; flex-direction: column; flex: 1; overflow: hidden; }}
        .header-container {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid #eaeaea; padding-bottom: 15px; flex-shrink: 0; }}
        h1 {{ font-size: 24px; color: #2c3e50; margin: 0; text-align: left; flex: 1; }}
        .btn-group {{ display: flex; gap: 10px; flex-wrap: wrap; justify-content: flex-end; }}
        .btn {{ color: white; border: none; padding: 8px 15px; font-size: 14px; font-weight: bold; border-radius: 6px; cursor: pointer; transition: background-color 0.3s ease; white-space: nowrap; }}
        .btn-blue {{ background-color: #3498db; }} .btn-blue:hover {{ background-color: #2980b9; }}
        .btn-green {{ background-color: #2ecc71; }} .btn-green:hover {{ background-color: #27ae60; }}
        .btn-gray {{ background-color: #7f8c8d; }} .btn-gray:hover {{ background-color: #95a5a6; }}
        .workspace {{ display: flex; gap: 20px; flex: 1; min-height: 0; }}
        .editor-pane {{ flex: 1; display: none; flex-direction: column; }}
        .editor-pane textarea {{ flex: 1; width: 100%; resize: none; font-family: 'Consolas', monospace; font-size: 14px; padding: 15px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; background-color: #fdfdfd; }}
        .diagram-pane {{ flex: 2; overflow: auto; display: flex; justify-content: center; align-items: flex-start; background: #fff; border: 1px dashed #ccc; border-radius: 6px; padding: 20px; }}
        #mermaid-container {{ width: 100%; text-align: center; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header-container">
            <h1>{title}</h1>
            <div class="btn-group">
                <button class="btn btn-gray" id="toggleEditorBtn" onclick="toggleEditor()">✏️ Редактировать</button>
                <button class="btn btn-blue" id="copyBtn" onclick="copyMermaidCode()">📋 Копировать код</button>
                <button class="btn btn-green" id="pngBtn" onclick="downloadPNG()">💾 Сохранить PNG</button>
            </div>
        </div>
        <div class="workspace">
            <div class="editor-pane" id="editorPane">
                <textarea id="rawCode" oninput="debounceUpdate()">{cleaned_code}</textarea>
            </div>
            <div class="diagram-pane" id="diagramPane">
                <div id="mermaid-container" class="mermaid">{cleaned_code}</div>
            </div>
        </div>
    </div>

    <script type="module">
        import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
        mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
        window.mermaidAPI = mermaid;
    </script>
    
    <script>
        let debounceTimer;
        
        // Задержка рендера при вводе кода, чтобы браузер не виснул
        function debounceUpdate() {{
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(updateDiagram, 500);
        }}

        // Обновление диаграммы в реальном времени
        async function updateDiagram() {{
            const code = document.getElementById('rawCode').value;
            const container = document.getElementById('mermaid-container');
            try {{
                const {{ svg }} = await window.mermaidAPI.render('mermaid-svg-live', code);
                container.innerHTML = svg;
            }} catch (e) {{
                // Игнорируем синтаксические ошибки в процессе печатания кода
                console.warn("Ожидание завершения ввода кода...");
            }}
        }}

        // Показать/Скрыть окно редактора кода
        function toggleEditor() {{
            const pane = document.getElementById('editorPane');
            const btn = document.getElementById('toggleEditorBtn');
            if (pane.style.display === 'flex') {{
                pane.style.display = 'none';
                btn.style.backgroundColor = '#7f8c8d';
            }} else {{
                pane.style.display = 'flex';
                btn.style.backgroundColor = '#e67e22'; // Оранжевый, когда активен
                updateDiagram();
            }}
        }}

        function copyMermaidCode() {{
            const rawCode = document.getElementById('rawCode').value;
            navigator.clipboard.writeText(rawCode).then(() => {{
                const btn = document.getElementById('copyBtn');
                const originalText = btn.innerText;
                btn.innerText = '✅ Скопировано!';
                setTimeout(() => {{ btn.innerText = originalText; }}, 2000);
            }});
        }}

        // Экспорт SVG в PNG высокого разрешения
        function downloadPNG() {{
            const svg = document.querySelector('#mermaid-container svg');
            if (!svg) {{ alert('Диаграмма не найдена!'); return; }}

            const bbox = svg.getBoundingClientRect();
            // Явно задаем размеры для корректного рендера на canvas (фикс для некоторых браузеров)
            svg.setAttribute('width', bbox.width);
            svg.setAttribute('height', bbox.height);

            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const data = new XMLSerializer().serializeToString(svg)
                .replace(/<br>/g, '<br/>')
                .replace(/<hr>/g, '<hr/>');

            const img = new Image();
            img.onload = function () {{
                // Масштабирование x2 для высокого качества (Retina)
                const scale = 2;
                canvas.width = bbox.width * scale;
                canvas.height = bbox.height * scale;
                ctx.scale(scale, scale);

                // Накладываем белый фон (по умолчанию PNG прозрачный)
                ctx.fillStyle = 'white';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0);

                // Скачивание
                const link = document.createElement('a');
                link.download = '{safe_title}.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
            }};
            img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(data);
        }}
    </script>
</body>
</html>"""

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        return f"Успешно! HTML-файл диаграммы сохранен. СКАЖИ ПОЛЬЗОВАТЕЛЮ ГОЛОСОМ (ТЕКСТОМ) ЧТО СХЕМА ГОТОВА И ОБЯЗАТЕЛЬНО ВЫВЕДИ ЭТУ ССЫЛКУ: [Из файла: {{filename}}]"
    except Exception as e:
        return f"Ошибка генерации диаграммы: {str(e)}"

def generate_yed_diagram(title: str, nodes: list, edges: list, app_instance=None) -> str:
    try:
        safe_title = re.sub(r'[\\/*?:"<>|]', "", title or "yEd_Diagram").replace(" ", "_").strip("._")
        if not safe_title:
            safe_title = "yEd_Diagram"
        filename = f"{safe_title}.graphml"

        if app_instance is not None:
            output_path = app_instance.ask_save_path_sync(filename, ext=".graphml")
            if not output_path:
                return "Сохранение диаграммы отменено пользователем."
        else:
            output_dir = os.path.join(get_base_path(), "Созданные_Документы", "Схемы")
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, filename)

        graphml_ns = "http://graphml.graphdrawing.org/xmlns"
        y_ns = "http://www.yworks.com/xml/graphml"
        xsi_ns = "http://www.w3.org/2001/XMLSchema-instance"

        ET.register_namespace("", graphml_ns)
        ET.register_namespace("y", y_ns)
        ET.register_namespace("xsi", xsi_ns)

        graphml = ET.Element(
            f"{{{graphml_ns}}}graphml",
            {
                f"{{{xsi_ns}}}schemaLocation": "http://graphml.graphdrawing.org/xmlns http://www.yworks.com/xml/schema/graphml/1.1/ygraphml.xsd"
            }
        )

        ET.SubElement(graphml, f"{{{graphml_ns}}}key", {"for": "node", "id": "d6", "yfiles.type": "nodegraphics"})
        ET.SubElement(graphml, f"{{{graphml_ns}}}key", {"for": "edge", "id": "d10", "yfiles.type": "edgegraphics"})
        main_graph = ET.SubElement(graphml, f"{{{graphml_ns}}}graph", {"edgedefault": "directed", "id": "G"})

        shape_map = {
            "start": ("ellipse", "#C0C0C0", "80", "40", "line"),
            "end": ("ellipse", "#C0C0C0", "80", "40", "line"),
            "process": ("roundrectangle", "#E8EEF7", "120", "40", "line"),
            "decision": ("diamond", "#FFCC00", "100", "60", "line"),
            "document": ("note", "#FFF9C4", "120", "40", "line"),
            "database": ("cylinder", "#FFFFFF", "80", "60", "line"),
            "manual_input": ("trapezoid", "#E8EEF7", "120", "40", "line"),
            "actor": ("rectangle", "#E0E0E0", "120", "40", "line"),
            "routing": ("ellipse", "#FF8C00", "15", "15", "line"),
            "idef_node": ("rectangle", "#F5F5F5", "120", "40", "dashed")
        }

        step = {"value": 0}

        def add_shape_node(parent_graph, node_obj):
            node_id = str(node_obj.get("id", "")).strip()
            if not node_id:
                return
            label = str(node_obj.get("label", "")).strip()
            shape_type = str(node_obj.get("shape", "process")).strip() or "process"

            x = str((step["value"] % 6) * 180)
            y = str((step["value"] // 6) * 120)
            step["value"] += 1

            if shape_type == "group":
                group_node = ET.SubElement(parent_graph, f"{{{graphml_ns}}}node", {"id": node_id, "yfiles.foldertype": "group"})
                group_data = ET.SubElement(group_node, f"{{{graphml_ns}}}data", {"key": "d6"})
                proxy = ET.SubElement(group_data, f"{{{y_ns}}}ProxyAutoBoundsNode")
                realizers = ET.SubElement(proxy, f"{{{y_ns}}}Realizers", {"active": "0"})
                group_shape = ET.SubElement(realizers, f"{{{y_ns}}}GroupNode")
                ET.SubElement(group_shape, f"{{{y_ns}}}Geometry", {"height": "150.0", "width": "240.0", "x": x, "y": y})
                ET.SubElement(group_shape, f"{{{y_ns}}}Fill", {"color": "#F5F5F5", "transparent": "false"})
                ET.SubElement(group_shape, f"{{{y_ns}}}BorderStyle", {"color": "#000000", "type": "dashed", "width": "1.0"})
                group_label = ET.SubElement(group_shape, f"{{{y_ns}}}NodeLabel", {
                    "alignment": "center",
                    "backgroundColor": "#EBEBEB",
                    "modelName": "internal",
                    "modelPosition": "t"
                })
                group_label.text = label
                ET.SubElement(group_shape, f"{{{y_ns}}}Shape", {"type": "roundrectangle"})
                ET.SubElement(group_shape, f"{{{y_ns}}}State", {"closed": "false", "innerGraphDisplayEnabled": "false"})
                ET.SubElement(group_shape, f"{{{y_ns}}}Insets", {"bottom": "15", "left": "15", "right": "15", "top": "15"})

                inner_graph = ET.SubElement(group_node, f"{{{graphml_ns}}}graph", {"edgedefault": "directed", "id": f"{node_id}:"})
                for child in node_obj.get("nodes", []) or []:
                    add_shape_node(inner_graph, child)
                return

            shape, fill_color, width, height, border_type = shape_map.get(shape_type, shape_map["process"])
            final_label = "" if shape_type == "routing" else label

            node_el = ET.SubElement(parent_graph, f"{{{graphml_ns}}}node", {"id": node_id})
            data_el = ET.SubElement(node_el, f"{{{graphml_ns}}}data", {"key": "d6"})
            shape_node = ET.SubElement(data_el, f"{{{y_ns}}}ShapeNode")
            ET.SubElement(shape_node, f"{{{y_ns}}}Geometry", {"height": height, "width": width, "x": x, "y": y})
            ET.SubElement(shape_node, f"{{{y_ns}}}Fill", {"color": fill_color, "transparent": "false"})
            ET.SubElement(shape_node, f"{{{y_ns}}}BorderStyle", {"color": "#000000", "type": border_type, "width": "1.0"})
            node_label = ET.SubElement(shape_node, f"{{{y_ns}}}NodeLabel")
            node_label.text = final_label
            ET.SubElement(shape_node, f"{{{y_ns}}}Shape", {"type": shape})

        for node in (nodes or []):
            add_shape_node(main_graph, node)

        for i, edge in enumerate(edges or []):
            source = str(edge.get("source", "")).strip()
            target = str(edge.get("target", "")).strip()
            if not source or not target:
                continue

            edge_el = ET.SubElement(main_graph, f"{{{graphml_ns}}}edge", {"id": f"e{i}", "source": source, "target": target})
            edge_data = ET.SubElement(edge_el, f"{{{graphml_ns}}}data", {"key": "d10"})
            poly_edge = ET.SubElement(edge_data, f"{{{y_ns}}}PolyLineEdge")
            ET.SubElement(poly_edge, f"{{{y_ns}}}Path", {"sx": "0.0", "sy": "0.0", "tx": "0.0", "ty": "0.0"})
            flow_type = str(edge.get("flow_type", "material")).strip() or "material"
            line_style = "line" if flow_type == "material" else "dashed"
            ET.SubElement(poly_edge, f"{{{y_ns}}}LineStyle", {"color": "#000000", "type": line_style, "width": "1.0"})
            ET.SubElement(poly_edge, f"{{{y_ns}}}Arrows", {"source": "none", "target": "standard"})
            edge_label = str(edge.get("label", "")).strip()
            if edge_label:
                edge_label_el = ET.SubElement(poly_edge, f"{{{y_ns}}}EdgeLabel")
                edge_label_el.text = edge_label

        tree = ET.ElementTree(graphml)
        tree.write(output_path, encoding="utf-8", xml_declaration=True)

        threading.Thread(target=sync_vector_db, daemon=True).start()
        return f"Успешно! GraphML-файл диаграммы сохранен. СКАЖИ ПОЛЬЗОВАТЕЛЮ ГОЛОСОМ (ТЕКСТОМ) ЧТО СХЕМА ГОТОВА И ОБЯЗАТЕЛЬНО ВЫВЕДИ ЭТУ ССЫЛКУ: [Из файла: {filename}]"
    except Exception as e:
        return f"Ошибка генерации yEd-диаграммы: {str(e)}"



# ИНДЕКСНОЕ РЕДАКТИРОВАНИЕ (Batch Index-Based Editing)
def apply_indexed_edits(filename, edits_list):
    """Применяет МАССИВ правок по индексам и сохраняет файл ОДИН раз"""
    target_file = find_target_file(filename)
    if not target_file:
        return f"Ошибка: Файл '{filename}' не найден в разрешенных директориях."

    try:
        doc = Document(target_file)
        all_paras = get_all_paragraphs(doc)
        
        # Применяем все правки в памяти
        for edit in edits_list:
            indices = edit.get("target_indices", [])
            new_text = edit.get("new_text", "").strip()
            if not indices: continue
            
            first_idx = min(indices)
            
            # 1. Зачеркиваем старое во всех указанных индексах
            for idx in indices:
                if idx < len(all_paras):
                    p = all_paras[idx]
                    old_text = p.text
                    for run in p.runs: run.text = "" # Очищаем
                    if old_text.strip():
                        del_run = p.add_run(old_text)
                        del_run.font.strike = True
                        del_run.font.color.rgb = RGBColor(255, 0, 0)
                        
            # 2. Вставляем новое (только в первый индекс блока)
            if first_idx < len(all_paras) and new_text and new_text.lower() not in ['delete', 'удалить']:
                p = all_paras[first_idx]
                p.add_run("\n[НОВАЯ РЕДАКЦИЯ]: ").font.bold = True
                new_run = p.add_run(new_text)
                new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        # Сохраняем результат ОДИН РАЗ после всех правок
        base, ext = os.path.splitext(target_file)
        output_path = f"{base}_Правки{ext}"
        doc.save(output_path)
        return f"Пакет правок успешно применен! Сохранено в {os.path.basename(output_path)}"
        
    except Exception as e:
        return f"Ошибка сохранения: {str(e)}"


# ==================== ИНСТРУМЕНТЫ АГЕНТА: УМНЫЙ EXCEL ====================
def smart_excel_search(filename, task_description, only_open=False, app_instance=None):
    """Инструмент 1: Умные Глаза (Поиск Топ-5 строк в Excel)"""
    target_file = find_target_file(filename)
    if not target_file:
        return f"Ошибка: Файл '{filename}' не найден в разрешенных директориях."

    try:
        global_settings = load_global_settings()
        local_settings = load_local_settings()
        if getattr(app_instance, "current_role", "guest") == "admin":
            current_model = local_settings.get("admin_model", "openai/gpt-4o-mini")
        else:
            current_model = local_settings.get("guest_model", "stepfun/step-3.5-flash:free")
        wb = openpyxl.load_workbook(target_file, data_only=True)
        sheet_names = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == 'visible']
        
        # ЭТАП 1: Разведка
        scout_data = {}
        for sheet in sheet_names:
            ws = wb[sheet]
            preview = []
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
                if any(cell is not None for cell in row): preview.append(f"Строка {i}: {row}")
            scout_data[sheet] = preview
            
        scout_prompt = "Ты Архитектор БД. Выбери 'target_sheet' и 'header_row_index' (строку с заголовками).\nВерни СТРОГО JSON: {\"target_sheet\": \"ИмяЛиста\", \"header_row_index\": 2}"
        scout_resp = get_llm_client().chat.completions.create(
            model=current_model, response_format={"type": "json_object"},
            messages=[{"role": "system", "content": scout_prompt}, {"role": "user", "content": f"Задача: {task_description}\n\nСтруктура:\n{json.dumps(scout_data, ensure_ascii=False)}"}]
        )
        scout_json = json.loads(re.search(r'\{.*\}', scout_resp.choices[0].message.content.strip(), re.DOTALL).group(0))
        target_sheet = scout_json.get("target_sheet", sheet_names[0])
        header_row_idx = int(scout_json.get("header_row_index", 1))
        
        ws = wb[target_sheet]
        
        headers_map = {}
        for cell in ws[header_row_idx]:
            if cell.value: headers_map[str(cell.value).replace('\n', ' ').strip()] = cell.column
        headers_list = list(headers_map.keys())

        sample_for_radar = []
        for r in range(header_row_idx + 1, min(header_row_idx + 15, ws.max_row + 1)):
            row_vals = {}
            is_empty = True
            for col_name, col_idx in headers_map.items():
                val = ws.cell(row=r, column=col_idx).value
                if val is not None and str(val).strip(): row_vals[col_name] = str(val).strip(); is_empty = False
            if not is_empty: sample_for_radar.append(row_vals)

        # ЭТАП 1.5: Колоночный Радар
        radar_prompt = (
            "Ты AI-Аналитик поиска. Определи правила поиска старой записи в таблице.\n"
            "СТРОГИЕ ПРАВИЛА:\n"
            "1. ИГНОРИРУЙ МЕТА-СЛОВА ('аудит', 'несоответствие', 'статус'). Ищи уникальную суть ('грязн', 'А06').\n"
            "2. ТИПИЗАЦИЯ: Колонки с '#', '№' или 'ID' - только для цифр/кодов.\n"
            "3. МУЛЬТИ-КОЛОНОЧНОСТЬ: ОБЯЗАТЕЛЬНО выбери МИНИМУМ 3 РАЗНЫЕ КОЛОНКИ для поиска (например: процесс, описание, причина). Если колонок мало, выбери все возможные. Это критически важно! Не ленись!\n\n"
            "Верни JSON: {\"search_rules\": [{\"column\": \"Точное Имя\", \"keywords\": [\"корень\"]}]}"
        )
        radar_resp = get_llm_client().chat.completions.create(
            model=current_model, response_format={"type": "json_object"},
            messages=[{"role": "system", "content": radar_prompt}, {"role": "user", "content": f"Задача: {task_description}\nКолонки: {headers_list}\nПримеры: {json.dumps(sample_for_radar[:3], ensure_ascii=False)}"}]
        )
        search_rules = json.loads(re.search(r'\{.*\}', radar_resp.choices[0].message.content.strip(), re.DOTALL).group(0)).get("search_rules", [])

        # Жесткий фильтр закрытых проблем (из настроек)
        status_col = global_settings.get("excel_status_col", "")
        closed_val = global_settings.get("excel_closed_val", "Выполнено").lower()

        scored_rows = []
        for r in range(header_row_idx + 1, ws.max_row + 1):
            row_dict = {"_ROW_INDEX_": r}
            is_empty = True
            for col_name, col_idx in headers_map.items():
                val = ws.cell(row=r, column=col_idx).value
                val_str = str(val).strip() if val is not None else ""
                row_dict[col_name] = val_str
                if val_str: is_empty = False
            
            if not is_empty:
                # Фильтр статуса
                if only_open and status_col in headers_map:
                    cell_status = row_dict.get(status_col, "").lower()
                    if closed_val in cell_status: continue # Пропускаем закрытые!

                if search_rules:
                    row_score = 0
                    for rule in search_rules:
                        col_to_search = rule.get("column")
                        kws = rule.get("keywords", [])
                        if col_to_search in headers_map and kws:
                            cell_val = row_dict.get(col_to_search, "").lower()
                            for kw in kws:
                                if kw.lower() in cell_val: row_score += 1
                    if row_score > 0:
                        scored_rows.append({"score": row_score, "data": row_dict})

        scored_rows.sort(key=lambda x: x["score"], reverse=True)
        targeted_sample = [item["data"] for item in scored_rows[:5]] # Берем Топ-5
        
        if not targeted_sample: return "Не найдено подходящих записей по вашему запросу."
        
        result_str = f"Найдено {len(targeted_sample)} кандидатов (Топ-5):\n"
        for row in targeted_sample: result_str += json.dumps(row, ensure_ascii=False) + "\n"
        return result_str

    except Exception as e: return f"Ошибка умного поиска Excel: {str(e)}"

def smart_excel_edit(filename, task_description, found_context_str, app_instance=None):
    """Инструмент 2: Умные Руки (Генерация JSON и Вставка в Excel)"""
    target_file = find_target_file(filename)
    if not target_file:
        return f"Ошибка: Файл '{filename}' не найден в разрешенных директориях."

    try:
        global_settings = load_global_settings()
        local_settings = load_local_settings()
        if getattr(app_instance, "current_role", "guest") == "admin":
            current_model = local_settings.get("admin_model", "openai/gpt-4o-mini")
        else:
            current_model = local_settings.get("guest_model", "stepfun/step-3.5-flash:free")
        
        base, ext = os.path.splitext(target_file)
        out_path = f"{base}_Правки{ext}"
        shutil.copy2(target_file, out_path)
        
        wb = openpyxl.load_workbook(out_path)
        sheet_names = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == 'visible']
        
        scout_data = {s: [f"Строка {i}: {row}" for i, row in enumerate(wb[s].iter_rows(min_row=1, max_row=5, values_only=True), 1)] for s in sheet_names}
        scout_resp = get_llm_client().chat.completions.create(
            model=current_model, response_format={"type": "json_object"},
            messages=[{"role": "system", "content": "Верни JSON: {\"target_sheet\": \"Имя\", \"header_row_index\": 2}"}, 
                      {"role": "user", "content": f"Задача: {task_description}\nСтруктура: {json.dumps(scout_data, ensure_ascii=False)}"}]
        )
        scout_json = json.loads(re.search(r'\{.*\}', scout_resp.choices[0].message.content.strip(), re.DOTALL).group(0))
        target_sheet = scout_json.get("target_sheet", sheet_names[0])
        header_row_idx = int(scout_json.get("header_row_index", 1))
        
        ws = wb[target_sheet]
        
        headers_map = {}
        headers_info = {}
        for cell in ws[header_row_idx]:
            if cell.value:
                col_name = str(cell.value).replace('\n', ' ').strip()
                headers_map[col_name] = cell.column
                comment = cell.comment.text if cell.comment else ""
                headers_info[col_name] = {"comment": comment.strip()} if comment else {}
                
        last_15_rows = []
        for r in range(max(header_row_idx + 1, ws.max_row - 14), ws.max_row + 1):
            row_dict = {}
            for c_name, c_idx in headers_map.items():
                val = ws.cell(row=r, column=c_idx).value
                if val is not None: row_dict[c_name] = str(val).strip()
            if row_dict: last_15_rows.append(row_dict)

        gen_prompt = (
            "Ты Аналитик паттернов СМК.\n"
            "ПРАВИЛА:\n"
            "1. Изучи 'НАЙДЕННЫЕ СТРОКИ'. Если обновляешь, используй точный '_ROW_INDEX_'. Возвращай ТОЛЬКО измененные колонки!\n"
            "2. Для новой записи используй '_ROW_INDEX_': 'new'.\n"
            "3. Продолжай паттерны нумерации из 'ПОСЛЕДНИХ СТРОК'.\n"
            f"4. ВАЖНО: Если меняешь статус, СТРОГО используй значения из системы. Открыто = '{global_settings.get('excel_open_val', 'Открыто')}', Закрыто = '{global_settings.get('excel_closed_val', 'Выполнено')}'.\n"
            "ВЕРНИ СТРОГО JSON:\n"
            '{"rows": [{"_ROW_INDEX_": "new", "Колонка": "Знач"}, {"_ROW_INDEX_": 111, "Статус": "Выполнено"}]}'
        )
        
        user_prompt = f"Задача: {task_description}\nКолонки и Примечания: {json.dumps(headers_info, ensure_ascii=False)}\n"
        if found_context_str: user_prompt += f"НАЙДЕННЫЕ СТРОКИ ДЛЯ ОБНОВЛЕНИЯ:\n{found_context_str}\n"
        user_prompt += f"ПОСЛЕДНИЕ СТРОКИ (Стиль): {json.dumps(last_15_rows, ensure_ascii=False)}"

        gen_resp = get_llm_client().chat.completions.create(
            model=current_model, response_format={"type": "json_object"},
            messages=[{"role": "system", "content": gen_prompt}, {"role": "user", "content": user_prompt}]
        )
        rows_to_process = json.loads(re.search(r'\{.*\}', gen_resp.choices[0].message.content.strip(), re.DOTALL).group(0)).get("rows", [])
        
        affected_rows = []
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for row_data in rows_to_process:
            row_idx_cmd = row_data.get("_ROW_INDEX_", "new")
            target_row = ws.max_row + 1 if (row_idx_cmd == "new" or not str(row_idx_cmd).isdigit()) else int(row_idx_cmd)
            affected_rows.append(str(target_row))
            comments_dict = row_data.get("_COMMENTS_", {})

            for col_name, value in row_data.items():
                if col_name in ["_ROW_INDEX_", "_COMMENTS_"]: continue
                col_idx = headers_map.get(col_name)
                if not col_idx:
                    for h_name, h_idx in headers_map.items():
                        if h_name.strip().lower() == col_name.strip().lower(): col_idx = h_idx; break
                
                if col_idx:
                    cell = ws.cell(row=target_row, column=col_idx)
                    cell.value = value
                    cell.fill = yellow_fill
                    if col_name in comments_dict:
                        cell.comment = Comment(text=str(comments_dict[col_name]), author="ИИ-Аналитик СМК")
                        
        wb.save(out_path)
        return f"Успех! Файл сохранен как: {os.path.basename(out_path)}. Изменены строки: {', '.join(affected_rows)}"

    except Exception as e: return f"Ошибка умного редактирования: {str(e)}"


# ==================== ИНСТРУМЕНТЫ АГЕНТА: OUTLOOK ====================
def draft_email_tool(to_name, subject, html_body):
    """Создать черновик письма в Outlook"""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_name if to_name else "Укажите email"
        mail.Subject = subject if subject else "Без темы"
        mail.HTMLBody = html_body if html_body else "<p>Текст письма...</p>"
        mail.Display()  # ТОЛЬКО Display! Никаких .Send()!
        return "Черновик письма успешно открыт в Outlook. Ожидает отправки пользователем."
    except Exception as e:
        return f"Ошибка подключения к Outlook: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

def draft_meeting_tool(to_name, subject, body, duration_minutes=60):
    """Создать черновик встречи в Outlook"""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        appt = outlook.CreateItem(1)  # 1 = olAppointmentItem
        appt.MeetingStatus = 1  # 1 = olMeeting
        appt.RequiredAttendees = to_name if to_name else "Укажите участников"
        appt.Subject = subject if subject else "Без темы"
        appt.Body = body if body else "Повестка встречи..." # Строго Body (без HTML)
        appt.Duration = duration_minutes
        appt.Display()  # ТОЛЬКО Display!
        return f"Черновик встречи ({duration_minutes} мин) успешно открыт в Outlook."
    except Exception as e:
        return f"Ошибка подключения к Outlook: {str(e)}"
    finally:
        pythoncom.CoUninitialize()


def generate_document_from_template(template_filename, task_description, new_filename, app_instance=None):
    """Инструмент: Создает новый документ по образцу с помощью Smart Clone & Clean Replace."""
    # 1. Ищем файл-шаблон
    target_file = find_target_file(template_filename)
    if not target_file:
        return f"Ошибка: Файл '{template_filename}' не найден в разрешенных директориях."

    try:
        # 2. Читаем шаблон
        template_text, all_paras = read_docx_with_indices(target_file)
        if not template_text: return "Ошибка: Не удалось прочитать шаблон."

        # 3. Запрашиваем JSON у LLM
        system_prompt = (
            "Ты эксперт СМК. Создай новый документ из шаблона.\n"
            "Тебе дадут текст старого документа с номерами абзацев [в скобках].\n"
            "Найди все старые даты, процессы, ФИО и мусор, которые нужно изменить.\n"
            "ВАЖНО: Верни ответ СТРОГО в формате JSON-объекта с ключом 'edits':\n"
            '{"edits": [{"target_indices": [3, 4], "new_text": "Новый текст или delete"}]}\n'
        )
        
        local_settings = load_local_settings()
        if getattr(app_instance, "current_role", "guest") == "admin":
            current_model = local_settings.get("admin_model", "openai/gpt-4o-mini")
        else:
            current_model = local_settings.get("guest_model", "stepfun/step-3.5-flash:free")
        response = get_llm_client().chat.completions.create(
            model=current_model,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Задача: {task_description}\n\nТекст:\n{template_text}"}
            ]
        )
        
        ans = response.choices[0].message.content.strip()
        match = re.search(r'\{.*\}', ans, re.DOTALL)
        if match: ans = match.group(0)
        
        edits_list = json.loads(ans).get("edits", [])
        
        # 4. Smart Clone
        if not new_filename.endswith('.docx'): new_filename += '.docx'
        if app_instance is not None:
            output_path = app_instance.ask_save_path_sync(new_filename, ext=".docx")
            if not output_path:
                return "Сохранение документа отменено пользователем."
        else:
            out_dir = os.path.join(get_base_path(), "Созданные_Документы")
            os.makedirs(out_dir, exist_ok=True)
            output_path = os.path.join(out_dir, new_filename)
        shutil.copy2(target_file, output_path)
        
        # 5. Clean Replace
        doc = Document(output_path)
        target_paras = get_all_paragraphs(doc)

        for edit in edits_list:
            indices = edit.get("target_indices", [])
            new_text = edit.get("new_text", "").strip()
            if not indices: continue
            first_idx = min(indices)
            
            original_font_name, original_font_size, original_bold, original_italic = None, None, None, None
            if first_idx < len(target_paras) and len(target_paras[first_idx].runs) > 0:
                first_run = target_paras[first_idx].runs[0]
                original_font_name = first_run.font.name
                original_font_size = first_run.font.size
                original_bold = first_run.font.bold
                original_italic = first_run.font.italic

            for idx in indices:
                if idx < len(target_paras):
                    for run in target_paras[idx].runs: run.text = ""
            
            if first_idx < len(target_paras) and new_text.lower() not in ['delete', 'удалить']:
                new_run = target_paras[first_idx].add_run(new_text)
                if original_font_name is not None: new_run.font.name = original_font_name
                if original_font_size is not None: new_run.font.size = original_font_size
                if original_bold is not None: new_run.font.bold = original_bold
                if original_italic is not None: new_run.font.italic = original_italic
                new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                
        doc.save(output_path)
        
        # Синхронизация базы в фоне
        threading.Thread(target=sync_vector_db, daemon=True).start()
        
        return f"Успешно! Документ по шаблону создан и сохранен как: {output_path}"

    except Exception as e:
        return f"Ошибка при генерации документа: {str(e)}"


def generate_document_from_scratch(task_description, new_filename, reference_filename="", app_instance=None):
    """Инструмент: Разработать АБСОЛЮТНО НОВЫЙ документ с нуля."""
    try:
        ref_text = ""
        target_file = None
        
        if reference_filename:
            target_file = find_target_file(reference_filename)
            if not target_file:
                return f"Ошибка: Файл '{reference_filename}' не найден в разрешенных директориях."
            if os.path.exists(target_file):
                ref_text, _ = read_docx_with_indices(target_file)
                
        system_prompt = (
            "Ты Главный Методолог СМК (ISO 9001). Твоя цель - разработать АБСОЛЮТНО НОВЫЙ документ с нуля.\n"
            "Сгенерируй документ в строгом JSON формате.\n"
            "Ключ 'document' должен содержать массив объектов с ключами 'type' (тип блока) и 'text' (содержимое).\n"
            "Допустимые типы: 'h1' (Главный заголовок), 'h2' (Подзаголовок), 'p' (Обычный абзац), 'bullet' (Пункт списка).\n"
            "Пример:\n"
            '{"document": [{"type": "h1", "text": "Политика Качества"}, {"type": "p", "text": "Текст..."}, {"type": "bullet", "text": "Пункт 1"}]}\n'
        )
        
        user_prompt = f"Задача: {task_description}\n"
        if ref_text:
            user_prompt += f"\nДля понимания стиля компании, вот пример корпоративного документа (используй тональность, но не копируй слепо):\n{ref_text[:3000]}"
            
        local_settings = load_local_settings()
        if getattr(app_instance, "current_role", "guest") == "admin":
            current_model = local_settings.get("admin_model", "openai/gpt-4o-mini")
        else:
            current_model = local_settings.get("guest_model", "stepfun/step-3.5-flash:free")
        response = get_llm_client().chat.completions.create(
            model=current_model, response_format={"type": "json_object"},
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}]
        )
        
        ans = response.choices[0].message.content.strip()
        match = re.search(r'\{.*\}', ans, re.DOTALL)
        if match: ans = match.group(0)
        doc_data = json.loads(ans).get("document", [])
        
        if not new_filename.endswith('.docx'): new_filename += '.docx'
        if app_instance is not None:
            output_path = app_instance.ask_save_path_sync(new_filename, ext=".docx")
            if not output_path:
                return "Сохранение документа отменено пользователем."
        else:
            out_dir = os.path.join(get_base_path(), "Созданные_Документы")
            os.makedirs(out_dir, exist_ok=True)
            output_path = os.path.join(out_dir, new_filename)
        
        if target_file and os.path.exists(target_file):
            shutil.copy2(target_file, output_path)
            doc = Document(output_path)
            for element in doc.element.body:
                if element.tag.endswith(('p', 'tbl', 'sectPr')):
                    if not element.tag.endswith('sectPr'):
                        element.getparent().remove(element)
        else:
            doc = Document()
            
        for block in doc_data:
            b_type = block.get("type", "p")
            b_text = block.get("text", "")
            
            if b_type == "h1":
                try: p = doc.add_paragraph(style='Heading 1')
                except KeyError:
                    try: p = doc.add_paragraph(style='Заголовок 1')
                    except KeyError: p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(b_text)
                run.font.name = 'Times New Roman'; run.font.size = Pt(16); run.font.bold = True; run.font.color.rgb = RGBColor(0, 0, 0)
                
            elif b_type == "h2":
                try: p = doc.add_paragraph(style='Heading 2')
                except KeyError:
                    try: p = doc.add_paragraph(style='Заголовок 2')
                    except KeyError: p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(b_text)
                run.font.name = 'Times New Roman'; run.font.size = Pt(14); run.font.bold = True; run.font.color.rgb = RGBColor(0, 0, 0)
                
            elif b_type == "bullet":
                try: p = doc.add_paragraph(style='List Bullet')
                except KeyError:
                    try: p = doc.add_paragraph(style='Маркированный список')
                    except KeyError: p = doc.add_paragraph(f"• ")
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                run = p.add_run(b_text)
                run.font.name = 'Times New Roman'; run.font.size = Pt(12)
                
            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Pt(35)
                run = p.add_run(b_text)
                run.font.name = 'Times New Roman'; run.font.size = Pt(12)
                
        doc.save(output_path)
        threading.Thread(target=sync_vector_db, daemon=True).start()
        return f"Успешно! Документ с нуля разработан и сохранен: {output_path}"
    except Exception as e: return f"Ошибка при генерации с нуля: {str(e)}"


# ==================== ИНСТРУМЕНТЫ АГЕНТА: EXCEL С НУЛЯ ====================
def generate_excel_from_scratch(task_description, new_filename, app_instance=None):
    """Инструмент: Создает новую многостраничную таблицу Excel с нуля по описанию."""
    try:
        system_prompt = (
            "Ты Эксперт по бизнес-таблицам СМК. Сгенерируй структуру Excel таблицы по запросу.\n"
            "Если задача подразумевает разделение данных, создай несколько листов (sheets).\n"
            "Верни СТРОГО JSON-объект с ключом 'sheets'. Каждый элемент массива - это лист с ключами 'sheet_name', 'headers' и 'rows'.\n"
            "Пример:\n"
            "{\"sheets\": [{\"sheet_name\": \"Риски\", \"headers\": [\"№\", \"Риск\"], \"rows\": [[\"1\", \"Отказ\"]]}, {\"sheet_name\": \"Справочник\", \"headers\": [\"ID\", \"Значение\"], \"rows\": [[\"A1\", \"Сервер\"]]}]}"
        )
        
        local_settings = load_local_settings()
        if getattr(app_instance, "current_role", "guest") == "admin":
            current_model = local_settings.get("admin_model", "openai/gpt-4o-mini")
        else:
            current_model = local_settings.get("guest_model", "stepfun/step-3.5-flash:free")
        response = get_llm_client().chat.completions.create(
            model=current_model,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Задача: {task_description}"}
            ]
        )
        
        ans = response.choices[0].message.content.strip()
        match = re.search(r'\{.*\}', ans, re.DOTALL)
        if match: ans = match.group(0)
        
        data = json.loads(ans)
        
        # Поддержка старого и нового формата JSON
        if "sheets" in data:
            sheets_data = data["sheets"]
        elif "headers" in data and "rows" in data:
            sheets_data = [{"sheet_name": "Таблица", "headers": data["headers"], "rows": data["rows"]}]
        else:
            return "Ошибка: Модель не сгенерировала корректную структуру (отсутствует ключ sheets)."

        wb = openpyxl.Workbook()
        default_sheet = wb.active
        first_sheet = True
        
        # Стили
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for s_info in sheets_data:
            sheet_name = str(s_info.get("sheet_name", "Таблица"))
            # Очистка имени листа (Excel ограничивает имя 31 символом и запрещает спецсимволы)
            sheet_name = re.sub(r'[\\*?:/\[\]]', '', sheet_name)[:31]
            headers = s_info.get("headers", [])
            rows = s_info.get("rows", [])
            
            if not headers: continue # Пропускаем пустые листы
            
            if first_sheet:
                ws = default_sheet
                ws.title = sheet_name
                first_sheet = False
            else:
                ws = wb.create_sheet(title=sheet_name)
                
            ws.append(headers)
            # Форматируем шапку
            for col_idx, cell in enumerate(ws[1], 1):
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border
                cell.alignment = center_alignment

            # Заполняем данные
            for row_data in rows:
                ws.append(row_data)
                
            # Форматируем ячейки данных
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = left_alignment

            # Автоподбор ширины колонок
            for col_idx, col_cells in enumerate(ws.columns, 1):
                max_length = 0
                column_letter = get_column_letter(col_idx)
                for cell in col_cells:
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except: pass
                ws.column_dimensions[column_letter].width = min((max_length + 2), 50)
                
            ws.auto_filter.ref = ws.dimensions

        if first_sheet: # Если так и не создали ни одного валидного листа
            return "Ошибка: Не удалось сгенерировать ни одного листа с колонками."

        if not new_filename.endswith('.xlsx'): new_filename += '.xlsx'
        if app_instance is not None:
            output_path = app_instance.ask_save_path_sync(new_filename, ext=".xlsx")
            if not output_path:
                return "Сохранение таблицы отменено пользователем."
        else:
            out_dir = os.path.join(get_base_path(), "Созданные_Документы")
            os.makedirs(out_dir, exist_ok=True)
            output_path = os.path.join(out_dir, new_filename)
        
        wb.save(output_path)
        
        # Индексируем в фоне
        threading.Thread(target=sync_vector_db, daemon=True).start()
        
        return f"Успешно! Многостраничная таблица Excel сгенерирована и сохранена: {output_path}"

    except Exception as e:
        return f"Ошибка при генерации Excel: {str(e)}"


# ==================== НАСТРОЙКИ СИСТЕМЫ (CONFIG FILE) ====================
DEFAULT_LOCAL_SETTINGS = {
    "guest_model": "stepfun/step-3.5-flash:free",
    "admin_model": "openai/gpt-4o-mini",
    "model_history": [],
    "use_proxy": False,
    "proxy_host": "127.0.0.1",
    "proxy_port": "2080"
}

DEFAULT_GLOBAL_SETTINGS = {
    "vision_model": "openai/gpt-4o-mini",
    "secretary_model": "stepfun/step-3.5-flash:free",
    "embedding_model": "qwen/qwen3-embedding-8b",
    "audio_provider": "OpenRouter",
    "audio_model": "openai/gpt-4o-audio-preview",
    "audio_chunk_mins": 60,
    "audio_overlap_secs": 15,
    "indexed_folders": [],
    "exclude_keywords": ["архив", "not_index", "old", "черновик", "секретно"],
    "default_excel_file": "Журнал регистрации результатов аудитов.xlsx",
    "excel_status_col": "Отметка о выполнении мероприятия",
    "excel_open_val": "Открыто",
    "excel_closed_val": "Выполнено",
    "chroma_batch_size": 100
}

def load_local_settings():
    settings_path = os.path.join(get_local_path(), "local_settings.json")
    current_settings = DEFAULT_LOCAL_SETTINGS.copy()
    try:
        if os.path.exists(settings_path):
            with open(settings_path, "r", encoding="utf-8") as f:
                current_settings.update(json.load(f))
    except Exception:
        pass
    return current_settings

def save_local_settings(data):
    try:
        settings_path = os.path.join(get_local_path(), "local_settings.json")
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

def load_global_settings():
    settings_path = os.path.join(get_base_path(), "global_settings.json")
    current_settings = DEFAULT_GLOBAL_SETTINGS.copy()
    try:
        if os.path.exists(settings_path):
            with open(settings_path, "r", encoding="utf-8") as f:
                current_settings.update(json.load(f))
    except Exception:
        pass
    return current_settings

def save_global_settings(data):
    try:
        settings_path = os.path.join(get_base_path(), "global_settings.json")
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


# ==================== GUI ПРИЛОЖЕНИЕ ====================

APP_NAME = "ИИ-Агент СМК"
APP_VERSION = "v1.6.0 Enterprise"
APP_DEVELOPER = "Плаксунов В.Б."
APP_PHONE = "2166"
APP_DESCRIPTION = (
    "1.6.0 - Появилась возможность читать аудио файлы и транскрибировать их.\n"
    "Корпоративный ИИ-ассистент для Системы Менеджмента Качества (СМК).\n"
    "Приложение помогает анализировать документы, выполнять аудит,\n"
    "искать информацию по базе знаний и формировать рабочие материалы.\n"
    "Поддерживается работа с Word, Excel, PDF и схемами GraphML в едином интерфейсе."
)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} | Версия: {APP_VERSION}")
        self.geometry("900x650")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.current_settings = load_local_settings()
        self.global_settings = load_global_settings()
        self.current_role = "guest"
        os.makedirs(os.path.join(get_local_path(), "Sessions"), exist_ok=True)
        self.current_session_id = str(uuid.uuid4())
        self.session_title = "Новый диалог"
        self.save_path_event = threading.Event()
        self.save_path_result = None
        self.save_path_queue = queue.Queue(maxsize=1)
        self.free_models_list = ["stepfun/step-3.5-flash:free", "google/gemini-2.0-flash-exp:free"]
        threading.Thread(target=self.fetch_free_models, daemon=True).start()
        
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        # Делаем 4-ю строку резиновой, чтобы прижать нижние элементы
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="ИИ-Агент СМК", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.settings_button = ctk.CTkButton(self.sidebar_frame, text="Настройки", command=self.open_settings)
        self.settings_button.grid(row=1, column=0, padx=20, pady=10)
        
        self.clear_button = ctk.CTkButton(self.sidebar_frame, text="Очистить чат", command=self.clear_chat)
        self.clear_button.grid(row=2, column=0, padx=20, pady=10)

        self.btn_history = ctk.CTkButton(self.sidebar_frame, text="📚 История диалогов", command=self.open_history_window)
        self.btn_history.grid(row=3, column=0, padx=20, pady=10)
        
        self.sync_button = ctk.CTkButton(self.sidebar_frame, text="Синхронизировать базу", command=self.manual_sync)
        self.sync_button.grid(row=4, column=0, padx=20, pady=(10, 0), sticky="s")
        self.btn_sync = self.sync_button

        self.auth_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="🔑 Войти как Админ",
            command=self.prompt_auth,
            fg_color="#455A64",
            hover_color="#263238"
        )
        self.auth_btn.grid(row=5, column=0, padx=20, pady=(8, 0), sticky="s")
        self.update_ui_for_role()

        self.progress_bar = ctk.CTkProgressBar(self.sidebar_frame)
        self.progress_bar.grid(row=6, column=0, padx=20, pady=(8, 4), sticky="ew")
        self.progress_bar.set(0)

        self.file_progress_label = ctk.CTkLabel(self.sidebar_frame, text="Ожидание синхронизации", font=ctk.CTkFont(size=11))
        self.file_progress_label.grid(row=7, column=0, padx=20, pady=(0, 6), sticky="w")
        
        self.status_label = ctk.CTkLabel(self.sidebar_frame, text="Загрузка...", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=8, column=0, padx=20, pady=(5, 15))
        
        self.chat_frame = ctk.CTkFrame(self)
        self.chat_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.chat_frame.grid_columnconfigure(0, weight=1)
        self.chat_frame.grid_rowconfigure(0, weight=1)
        
        self.chat_textbox = ctk.CTkTextbox(self.chat_frame, wrap="word")
        self.chat_textbox.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")
        self.chat_textbox.configure(state="disabled")

        self.link_map = {}
        self.link_counter = 0

        text_widget = self.chat_textbox._textbox
        text_widget.tag_config("hide", elide=True)
        text_widget.tag_config("bold", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
        text_widget.tag_config("h1", font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"), foreground="#FF8C00")
        text_widget.tag_config("h2", font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"), foreground="#00BFFF")
        text_widget.tag_config("h3", font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), foreground="#4FC3F7")
        # Убрали фон и цвет текста у таблицы для адаптивности под темы
        text_widget.tag_config("table", font=ctk.CTkFont(family="Consolas", size=13), wrap="none", justify="center")
        text_widget.tag_config("hr", foreground="#555555")
        text_widget.tag_config("hyperlink", foreground="#1f538d", underline=True)
        text_widget.tag_bind("hyperlink", "<Enter>", lambda e: text_widget.config(cursor="hand2"))
        text_widget.tag_bind("hyperlink", "<Leave>", lambda e: text_widget.config(cursor=""))
        
        # --- НОВОЕ: Стили бейджей (плашек) ---
        text_widget.tag_config("user_msg", background="#1F6AA5", foreground="white", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
        text_widget.tag_config("agent_msg", background="#555555", foreground="white", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
        text_widget.tag_bind("hyperlink", "<Button-1>", self.on_link_click)
        
        self.input_frame = ctk.CTkFrame(self.chat_frame)
        self.input_frame.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="ew")
        self.input_frame.grid_columnconfigure(0, weight=1)
        
        # Многострочное текстовое поле (высота 80px - это примерно 3-4 строки)
        self.input_entry = ctk.CTkTextbox(self.input_frame, font=ctk.CTkFont(size=14), height=80, wrap="word")
        self.input_entry.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="ew")

        # Обработка Enter (Отправка) и Shift+Enter (Перенос строки)
        def enter_pressed(event):
            # Если нажат Shift (состояние включает 0x0001) - разрешаем стандартный перенос
            if event.state & 0x0001:
                return None
            else:
                self.send_message()
                return "break" # Блокируем добавление новой строки при отправке

        self.input_entry.bind("<Return>", enter_pressed)
        
        self.send_button = ctk.CTkButton(self.input_frame, text="Отправить", width=100, command=self.send_message)
        self.send_button.grid(row=0, column=1, padx=(0, 0), pady=10)
        
        self.chat_history = []
        self.load_history()
        
        def init_db_thread():
            try:
                _, count = sync_vector_db(self)
                self.after(0, lambda: self.status_label.configure(text=f"База готова (чанков: {count})"))
            except Exception as e:
                self.after(0, lambda: self.status_label.configure(text=f"Ошибка БД: {e}"))
        threading.Thread(target=init_db_thread, daemon=True).start()

    def fetch_free_models(self):
        try:
            response = requests.get("https://openrouter.ai/api/v1/models", timeout=10)
            response.raise_for_status()
            data = response.json()
            models = data.get("data", []) if isinstance(data, dict) else []

            free_models = []
            for model in models:
                pricing = model.get("pricing", {}) if isinstance(model, dict) else {}
                if str(pricing.get("prompt", "")).strip() == "0" and str(pricing.get("completion", "")).strip() == "0":
                    model_id = model.get("id")
                    if model_id:
                        free_models.append(model_id)

            if free_models:
                self.free_models_list = sorted(set(free_models))
        except Exception:
            pass

    def prompt_auth(self):
        if self.current_role == "admin":
            self.current_role = "guest"
            self.update_ui_for_role()
            self.append_to_chat("\n[Система: Режим администратора отключен. Текущая роль: Guest.]\n")
            return

        password_dialog = ctk.CTkInputDialog(text="Введите пароль администратора:", title="Авторизация")
        entered_password = password_dialog.get_input() if password_dialog else None

        if entered_password == get_vault_data().get("admin_password", "admin"):
            self.current_role = "admin"
            self.update_ui_for_role()
            self.append_to_chat("\n[Система: Успешная авторизация. Текущая роль: Admin.]\n")
        else:
            self.append_to_chat("\n[Система: Неверный пароль. Доступ Admin отклонен.]\n")

    def update_ui_for_role(self):
        is_admin = self.current_role == "admin"
        if hasattr(self, "btn_sync"):
            self.btn_sync.configure(state="normal" if is_admin else "disabled")
        if hasattr(self, "btn_history"):
            if is_admin:
                self.btn_history.grid()
            else:
                self.btn_history.grid_remove()
        if hasattr(self, "auth_btn"):
            self.auth_btn.configure(text="🔒 Выйти (Админ)" if is_admin else "🔑 Войти как Админ")

    def ask_save_path_sync(self, suggested_filename, ext=".docx"):
        self.save_path_result = None
        self.save_path_event.clear()
        self.after(0, self._show_save_dialog, suggested_filename, ext)
        self.save_path_event.wait()
        return self.save_path_result

    def _show_save_dialog(self, suggested_filename, ext):
        try:
            normalized_ext = ext if str(ext).startswith(".") else f".{ext}"
            if not str(suggested_filename).lower().endswith(normalized_ext.lower()):
                suggested_filename = f"{suggested_filename}{normalized_ext}"

            out_dir = os.path.abspath("SMK_Docs/Созданные_Документы")
            os.makedirs(out_dir, exist_ok=True)
            selected_path = filedialog.asksaveasfilename(
                title="Сохранить файл как",
                initialdir=out_dir,
                initialfile=suggested_filename,
                defaultextension=normalized_ext,
                filetypes=[(f"*{normalized_ext}", f"*{normalized_ext}"), ("Все файлы", "*.*")]
            )
            self.save_path_result = selected_path if selected_path else None
        finally:
            self.save_path_event.set()

    def update_progress_ui(self, progress, filename):
        self.progress_bar.set(progress)
        if filename == "Синхронизация завершена":
            self.file_progress_label.configure(text="Синхронизация завершена")
        else:
            self.file_progress_label.configure(text=f"Текущий файл: {filename}")
        self.update_idletasks()

    def append_to_chat(self, text, tags=None):
        self.chat_textbox.configure(state="normal")
        if tags:
            self.chat_textbox.insert("end", text, tags)
        else:
            self.chat_textbox.insert("end", text)
        self.chat_textbox.see("end")
        self.chat_textbox.configure(state="disabled")

    def generate_unicode_table(self, raw_table, max_chars=100):
        lines = raw_table.strip().split('\n')
        parsed_rows = []

        for line in lines:
            if re.match(r'^[ \t]*\|?[-: |]+\|?[ \t]*$', line):
                continue
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if cells:
                clean_cells = [cell.replace('**', '') for cell in cells]
                parsed_rows.append(clean_cells)

        if not parsed_rows:
            return raw_table

        cols_count = max(len(r) for r in parsed_rows)
        col_widths = [0] * cols_count

        for row in parsed_rows:
            for i, cell in enumerate(row):
                if i < cols_count:
                    col_widths[i] = max(col_widths[i], len(cell))

        total_width = sum(col_widths) + cols_count * 3 + 1

        if total_width > max_chars:
            target_avg = max(5, (max_chars - cols_count * 3 - 1) // cols_count)
            allocated = [min(w, target_avg) for w in col_widths]
            remaining = (max_chars - cols_count * 3 - 1) - sum(allocated)

            while remaining > 0:
                added = False
                for i in range(cols_count):
                    if allocated[i] < col_widths[i] and remaining > 0:
                        allocated[i] += 1
                        remaining -= 1
                        added = True
                if not added:
                    break
            col_widths = allocated
        elif total_width < max_chars:
            remaining = max_chars - total_width
            idx = 0
            while remaining > 0:
                col_widths[idx % cols_count] += 1
                remaining -= 1
                idx += 1

        col_widths = [max(3, w) for w in col_widths]

        def build_separator(left, mid, right, fill):
            return left + mid.join(fill * w for w in col_widths) + right

        top_border = build_separator('┌─', '─┬─', '─┐', '─')
        mid_border = build_separator('├─', '─┼─', '─┤', '─')
        bot_border = build_separator('└─', '─┴─', '─┘', '─')

        formatted_lines = [top_border]

        for r_idx, row in enumerate(parsed_rows):
            while len(row) < cols_count:
                row.append("")

            wrapped_cells = [
                textwrap.wrap(cell, width=col_widths[i]) if col_widths[i] > 0 else [""]
                for i, cell in enumerate(row)
            ]
            max_lines = max((len(c) for c in wrapped_cells), default=1)

            for line_idx in range(max_lines):
                row_str = "│"
                for col_idx in range(cols_count):
                    cell_lines = wrapped_cells[col_idx]
                    text = cell_lines[line_idx] if line_idx < len(cell_lines) else ""

                    if r_idx == 0:
                        row_str += " " + text.center(col_widths[col_idx]) + " │"
                    else:
                        row_str += " " + text.ljust(col_widths[col_idx]) + " │"

                formatted_lines.append(row_str)

            if r_idx < len(parsed_rows) - 1:
                formatted_lines.append(mid_border)

        formatted_lines.append(bot_border)
        return "\n" + "\n".join(formatted_lines) + "\n"

    def apply_markdown(self, start_index):
        self.chat_textbox.configure(state="normal")
        text_widget = self.chat_textbox._textbox
        end_index = self.chat_textbox.index("end-1c")

        pixel_width = text_widget.winfo_width()
        calculated_max_chars = max(50, (pixel_width - 40) // 8)

        raw_text = text_widget.get(start_index, end_index)
        table_matches = list(re.finditer(r'(^[ \t]*\|.*\|[ \t]*(\n|$))+', raw_text, re.MULTILINE))

        for match in reversed(table_matches):
            raw_table = match.group(0)
            unicode_table = self.generate_unicode_table(raw_table, max_chars=calculated_max_chars)

            m_start, m_end = match.start(), match.end()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"

            text_widget.delete(tk_start, tk_end)
            text_widget.insert(tk_start, unicode_table)

            new_tk_end = f"{tk_start} + {len(unicode_table)} chars"
            text_widget.tag_add("table", tk_start, new_tk_end)

        end_index = self.chat_textbox.index("end-1c")
        raw_text = text_widget.get(start_index, end_index)

        for match in re.finditer(r'^(#{1,3})\s+(.*?)$', raw_text, re.MULTILINE):
            hashes = match.group(1)
            level = len(hashes)
            m_start, m_end = match.start(), match.end()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"
            hash_end = f"{start_index} + {m_start + level + 1} chars"
            text_widget.tag_add(f"h{level}", tk_start, tk_end)
            text_widget.tag_add("hide", tk_start, hash_end)

        for match in re.finditer(r'\*\*(.*?)\*\*', raw_text):
            m_start, m_end = match.start(), match.end()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"
            tk_inner_start = f"{start_index} + {m_start + 2} chars"
            tk_inner_end = f"{start_index} + {m_end - 2} chars"
            text_widget.tag_add("bold", tk_inner_start, tk_inner_end)
            text_widget.tag_add("hide", tk_start, tk_inner_start)
            text_widget.tag_add("hide", tk_inner_end, tk_end)

        for match in re.finditer(r'^---$', raw_text, re.MULTILINE):
            m_start, m_end = match.start(), match.end()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"
            text_widget.tag_add("hr", tk_start, tk_end)

        # 4. Ссылки (Улучшенная "всеядная" регулярка)
        for match in re.finditer(r'\[(?:Из файла|Файл)[:\s]*([^\]]+)\]', raw_text, re.IGNORECASE):
            m_start, m_end = match.start(), match.end()
            filename = match.group(1).strip()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"
            link_tag = f"link_{self.link_counter}"
            self.link_map[link_tag] = filename
            self.link_counter += 1
            text_widget.tag_add("hyperlink", tk_start, tk_end)
            text_widget.tag_add(link_tag, tk_start, tk_end)

        # 5. Веб-ссылки (http/https)
        for match in re.finditer(r'(https?://[^\s\]\)]+)', raw_text):
            m_start, m_end = match.start(), match.end()
            url = match.group(1).strip()
            tk_start = f"{start_index} + {m_start} chars"
            tk_end = f"{start_index} + {m_end} chars"
            link_tag = f"weblink_{self.link_counter}"
            self.link_map[link_tag] = url
            self.link_counter += 1
            text_widget.tag_add("hyperlink", tk_start, tk_end)
            text_widget.tag_add(link_tag, tk_start, tk_end)

        self.chat_textbox.configure(state="disabled")

    def on_link_click(self, event):
        text_widget = self.chat_textbox._textbox
        index = text_widget.index(f"@{event.x},{event.y}")
        tags = text_widget.tag_names(index)
        
        for tag in tags:
            if tag.startswith("link_"):
                filename = self.link_map.get(tag)
                if filename:
                    # Используем наш новый универсальный локатор!
                    target_file = find_target_file(filename)
                    
                    if target_file and os.path.exists(target_file):
                        if self.current_role == "guest":
                            base_name = os.path.basename(target_file)
                            safe_filename = f"СМК_Чтение_{base_name}"
                            safe_path = os.path.join(tempfile.gettempdir(), safe_filename)
                            shutil.copy2(target_file, safe_path)
                            self.append_to_chat(f"\n[Система: Guest-режим. Открываем безопасную копию: '{safe_path}']\n")
                            os.startfile(os.path.abspath(safe_path))
                        else:
                            self.append_to_chat(f"\n[Система: Admin-режим. Открываем оригинал файла '{filename}']\n")
                            os.startfile(os.path.abspath(target_file))
                    else:
                        self.append_to_chat(f"\n[Система: Файл '{filename}' не найден в разрешенных директориях]\n")
                break
            elif tag.startswith("weblink_"):
                url = self.link_map.get(tag)
                if url:
                    webbrowser.open(url)
                break
    
    def load_history(self):
        try:
            history_path = os.path.join(get_local_path(), "chat_history.json")
            if os.path.exists(history_path):
                with open(history_path, 'r', encoding='utf-8') as f: self.chat_history = json.load(f)
        except: pass
    
    def save_history(self):
        try:
            history_path = os.path.join(get_local_path(), "chat_history.json")
            with open(history_path, 'w', encoding='utf-8') as f: json.dump(self.chat_history[-40:], f, ensure_ascii=False, indent=2)
        except: pass

    def generate_session_title_background(self, first_prompt):
        try:
            secretary_model = self.global_settings.get("secretary_model", "openai/gpt-4o-mini") or "openai/gpt-4o-mini"
            response = get_llm_client().chat.completions.create(
                model=secretary_model,
                messages=[
                    {
                        "role": "system",
                        "content": "Сформируй краткий заголовок сессии чата на 4-5 слов. Ответь только заголовком без кавычек и пояснений."
                    },
                    {
                        "role": "user",
                        "content": first_prompt
                    }
                ]
            )
            title = (response.choices[0].message.content or "").strip()
            title = re.sub(r'[\\/*?:"<>|]', "", title)
            if title:
                self.session_title = title
                self.save_current_session()
        except Exception as e:
            print(f"Ошибка фонового нейминга: {e}")

    def save_current_session(self):
        """Автосохранение текущего состояния чата в JSON."""
        # Блокировка: Гости не оставляют следов на диске
        if getattr(self, "current_role", "guest") != "admin":
            return

        if not self.chat_history:
            return
        try:
            sessions_dir = os.path.join(get_local_path(), "Sessions")
            os.makedirs(sessions_dir, exist_ok=True)
            file_path = os.path.join(sessions_dir, f"{self.current_session_id}.json")
            display_text = self.chat_textbox._textbox.get("1.0", "end-1c")
            payload = {
                "session_id": self.current_session_id,
                "title": self.session_title,
                "timestamp": datetime.now().isoformat(),
                "chat_history": self.chat_history,
                "display_text": display_text
            }
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения сессии: {e}")

    def load_session(self, session_id, window_to_close=None):
        file_path = os.path.join(get_local_path(), "Sessions", f"{session_id}.json")
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.current_session_id = data.get("session_id", session_id)
            self.session_title = data.get("title", "Новый диалог")
            self.chat_history = data.get("chat_history", [])
            display_text = data.get("display_text", "")

            self.chat_textbox.configure(state="normal")
            self.chat_textbox.delete("1.0", "end")
            self.chat_textbox.insert("1.0", display_text)
            self.chat_textbox.configure(state="disabled")
            self.apply_markdown("1.0")

            if window_to_close is not None:
                window_to_close.destroy()

            self.append_to_chat(f"\n[Система: Загружена сессия '{self.session_title}']\n")
        except Exception as e:
            self.append_to_chat(f"\n[Система: Ошибка загрузки сессии: {e}]\n")

    def open_history_window(self):
        history_window = ctk.CTkToplevel(self)
        history_window.title("История диалогов")
        history_window.geometry("600x400")
        history_window.transient(self)
        history_window.grab_set()

        scrollable = ctk.CTkScrollableFrame(history_window)
        scrollable.pack(fill="both", expand=True, padx=12, pady=12)

        session_files = sorted(
            glob.glob(os.path.join(get_local_path(), "Sessions", "*.json")),
            key=os.path.getmtime,
            reverse=True
        )
        if not session_files:
            ctk.CTkLabel(scrollable, text="Нет сохраненных сессий.").pack(pady=10)
            return

        def delete_session(file_path, sid, row_frame):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                self.append_to_chat(f"\n[Система: Ошибка удаления файла сессии: {e}]\n")
                return

            try:
                client = chromadb.PersistentClient(path=get_db_path())
                collection = client.get_or_create_collection(name="temp_chat_memory", embedding_function=get_cloud_ef())
                collection.delete(where={"session_id": sid})
            except Exception as e:
                print(f"Ошибка удаления из Chroma: {e}")

            row_frame.destroy()

            if sid == self.current_session_id:
                self.chat_textbox.configure(state="normal")
                self.chat_textbox.delete("1.0", "end")
                self.chat_textbox.configure(state="disabled")
                self.chat_history = []
                self.current_session_id = str(uuid.uuid4())
                self.session_title = "Новый диалог"

        for file_path in session_files:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                continue

            sid = data.get("session_id")
            title = data.get("title", "Новый диалог")
            timestamp = data.get("timestamp", "")

            row_frame = ctk.CTkFrame(scrollable)
            row_frame.pack(fill="x", padx=4, pady=4)

            ctk.CTkLabel(
                row_frame,
                text=f"{title}\n{timestamp}",
                anchor="w",
                justify="left"
            ).pack(side="left", fill="x", expand=True, padx=8, pady=8)

            ctk.CTkButton(
                row_frame,
                text="Загрузить",
                width=90,
                command=lambda session_id=sid: self.load_session(session_id, history_window)
            ).pack(side="right", padx=(6, 8), pady=8)

            ctk.CTkButton(
                row_frame,
                text="🗑",
                width=40,
                fg_color="#8E2A2A",
                hover_color="#6D1F1F",
                command=lambda fp=file_path, session_id=sid, rf=row_frame: delete_session(fp, session_id, rf)
            ).pack(side="right", padx=(0, 0), pady=8)

    def run_background_secretary(self, recent_messages):
        """Фоновый Секретарь СМК - анализирует диалог и запоминает новые факты"""
        try:
            # БЛОКИРОВКА: Фоновый секретарь работает только у Администратора!
            if getattr(self, "current_role", "guest") != "admin":
                return

            # Формируем контекст из последних сообщений
            context = "\n".join([f"{m.get('role', 'unknown')}: {m.get('content', '')[:200]}" for m in recent_messages])
            
            system_prompt = (
                "Ты фоновый Секретарь СМК. Твоя цель: проанализировать диалог и найти НОВЫЕ утвержденные факты или правила СМК "
                "(например: процессы переданы подрядчику, изменились стандарты). Игнорируй вопросы, гипотезы и обычный поиск. "
                "Верни СТРОГО JSON: {\"is_new_fact\": true/false, \"fact_text\": \"Полный текст для базы\", \"summary\": \"Краткая суть для лога в чат\"}."
            )
            
            response = get_llm_client().chat.completions.create(
                model=self.global_settings.get("secretary_model", "stepfun/step-3.5-flash:free"),
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Проанализируй этот диалог:\n{context}"}
                ],
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            
            if result.get("is_new_fact", False):
                fact_text = result.get("fact_text", "")
                summary = result.get("summary", "")
                if fact_text:
                    memorize_important_fact(fact_text)
                    msg = f"\n[🤫 Фоновый Секретарь: Запомнил новый факт СМК - {summary}]\n\n"
                    self.after(0, lambda: self.append_to_chat(msg))
        except Exception:
            # Отказоустойчивость: silently fail
            pass

    def clear_chat(self):
        self.chat_textbox.configure(state="normal")
        self.chat_textbox.delete("1.0", "end")
        self.chat_textbox.configure(state="disabled")
        self.chat_history = []
        self.current_session_id = str(uuid.uuid4())
        self.session_title = "Новый диалог"
        self.save_history()
        self.save_current_session()
        
        # --- НОВОЕ: Очистка временного архива ---
        try:
            client = chromadb.PersistentClient(path=get_db_path())
            collection = client.get_or_create_collection(name="temp_chat_memory", embedding_function=get_cloud_ef())
            collection.delete(where={"session_id": self.current_session_id})
            self.append_to_chat("\n[СИСТЕМА: Ваш личный архив диалога очищен]\n\n")
        except:
            pass # Если коллекции еще нет, игнорируем

    def manual_sync(self):
        self.status_label.configure(text="Синхронизация...")
        self.sync_button.configure(state="disabled")
        def do_sync():
            try:
                _, count = sync_vector_db(self)
                self.after(0, lambda: self.status_label.configure(text=f"База готова (чанков: {count})"))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self.status_label.configure(text=msg))
                print(f"Sync error: {error_msg}")
            finally:
                self.after(0, lambda: self.sync_button.configure(state="normal"))
        threading.Thread(target=do_sync, daemon=True).start()

    def open_settings(self):
        settings_window = ctk.CTkToplevel(self)
        settings_window.title("Настройки Агента СМК")
        settings_window.geometry("550x600") # Чуть увеличили высоту для поля пароля
        settings_window.transient(self)
        settings_window.grab_set()

        tabview = ctk.CTkTabview(settings_window)
        tabview.pack(padx=20, pady=10, fill="both", expand=True)

        tab_models = tabview.add("Модели")
        tab_excludes = tabview.add("Исключения")
        tab_folders = tabview.add("Папки")
        tab_about = tabview.add("О программе ℹ️")
        
        # Флаг для удобного разделения прав
        is_admin = (self.current_role == "admin")
        tab_security = tabview.add("Безопасность 🔒") if is_admin else None
        
        # --- ВКЛАДКА: АУДИО И СЕТЬ (ТОЛЬКО ДЛЯ АДМИНА) ---
        tab_audio_net = tabview.add("Аудио и Сеть") if is_admin else None

        # --- ВКЛАДКА 1: МОДЕЛИ ---
        ctk.CTkLabel(tab_models, text="ID Модели (OpenRouter):").pack(pady=(10, 0))
        
        if is_admin:
            # АДМИН: Редактируемый список с историей Топ-10
            history = self.current_settings.get("model_history", [])
            model_entry = ctk.CTkComboBox(tab_models, width=450, values=history)
            model_entry.set(self.current_settings.get("admin_model", "openai/gpt-4o-mini"))
        else:
            # ГОСТЬ: Только чтение, список бесплатных моделей
            model_entry = ctk.CTkComboBox(tab_models, width=450, values=self.free_models_list, state="readonly")
            model_entry.set(self.current_settings.get("guest_model", "stepfun/step-3.5-flash:free"))
        model_entry.pack(pady=5)

        ctk.CTkLabel(tab_models, text="Модель для Vision (OCR сканов и схем):").pack(pady=(10, 0))
        vision_entry = ctk.CTkEntry(tab_models, width=450)
        vision_entry.pack(pady=5)
        vision_entry.insert(0, self.global_settings.get("vision_model", "openai/gpt-4o-mini"))
        if not is_admin: vision_entry.configure(state="disabled", text_color="gray")

        ctk.CTkLabel(tab_models, text="Модель Фонового Секретаря:").pack(pady=(10, 0))
        secretary_entry = ctk.CTkEntry(tab_models, width=450)
        secretary_entry.pack(pady=5)
        secretary_entry.insert(0, self.global_settings.get("secretary_model", "openai/gpt-4o-mini"))
        if not is_admin: secretary_entry.configure(state="disabled", text_color="gray")

        ctk.CTkLabel(tab_models, text="Модель Эмбеддингов (нужен перезапуск):").pack(pady=(10, 0))
        embed_entry = ctk.CTkEntry(tab_models, width=450)
        embed_entry.pack(pady=5)
        embed_entry.insert(0, self.global_settings.get("embedding_model", "qwen/qwen3-embedding-8b"))
        if not is_admin: embed_entry.configure(state="disabled", text_color="gray")

        openrouter_entry = None
        groq_entry = None
        tavily_entry = None
        admin_pwd_entry = None
        if is_admin and tab_security is not None:
            vault_data = get_vault_data()

            ctk.CTkLabel(tab_security, text="OpenRouter API Key:").pack(pady=(10, 0))
            openrouter_entry = ctk.CTkEntry(tab_security, width=450, show="*")
            openrouter_entry.pack(pady=5)
            openrouter_entry.insert(0, vault_data.get("openrouter_key", ""))

            ctk.CTkLabel(tab_security, text="Groq API Key:").pack(pady=(10, 0))
            groq_entry = ctk.CTkEntry(tab_security, width=450, show="*")
            groq_entry.pack(pady=5)
            groq_entry.insert(0, vault_data.get("groq_key", ""))

            ctk.CTkLabel(tab_security, text="Tavily API Key:").pack(pady=(10, 0))
            tavily_entry = ctk.CTkEntry(tab_security, width=450, show="*")
            tavily_entry.pack(pady=5)
            tavily_entry.insert(0, vault_data.get("tavily_key", ""))

            ctk.CTkLabel(tab_security, text="Пароль администратора:").pack(pady=(10, 0))
            admin_pwd_entry = ctk.CTkEntry(tab_security, width=450, show="*")
            admin_pwd_entry.pack(pady=5)
            admin_pwd_entry.insert(0, vault_data.get("admin_password", "admin"))

        # --- ВКЛАДКА 2: ИСКЛЮЧЕНИЯ ---
        ctk.CTkLabel(tab_excludes, text="Слова-исключения для папок/файлов (через запятую):").pack(pady=(10, 5))
        excludes_entry = ctk.CTkTextbox(tab_excludes, width=450, height=150, wrap="word")
        excludes_entry.pack(pady=5)
        excludes_text = ", ".join(self.global_settings.get("exclude_keywords", ["архив", "not_index"]))
        excludes_entry.insert("1.0", excludes_text)
        if not is_admin: excludes_entry.configure(state="disabled", text_color="gray")

        # --- ВКЛАДКА 3: ИНТЕРАКТИВНЫЕ ПАПКИ ---
        temp_folders = self.global_settings.get("indexed_folders", ["./SMK_Docs", "./Memory"]).copy()

        def render_folders():
            for widget in folders_scroll.winfo_children():
                widget.destroy()
                
            for f_path in temp_folders:
                row = ctk.CTkFrame(folders_scroll, fg_color="transparent")
                row.pack(fill="x", pady=3)
                
                lbl = ctk.CTkLabel(row, text=f_path, anchor="w")
                lbl.pack(side="left", padx=5, fill="x", expand=True)
                
                # Кнопка удаления только для Админа
                if is_admin:
                    btn = ctk.CTkButton(row, text="−", width=30, height=24, fg_color="#D32F2F", hover_color="#B71C1C",
                                        command=lambda p=f_path: remove_folder(p))
                    btn.pack(side="right", padx=5)

        def add_folder():
            folder_path = ctk.filedialog.askdirectory(title="Выберите папку для базы СМК")
            if folder_path and folder_path not in temp_folders:
                temp_folders.append(folder_path)
                render_folders()

        def remove_folder(path_to_remove):
            if path_to_remove in temp_folders:
                temp_folders.remove(path_to_remove)
                render_folders()

        # Кнопка добавления только для Админа
        if is_admin:
            add_btn = ctk.CTkButton(tab_folders, text="+ Добавить папку", command=add_folder)
            add_btn.pack(pady=(10, 5))

        folders_scroll = ctk.CTkScrollableFrame(tab_folders, width=450, height=250)
        folders_scroll.pack(pady=5, fill="both", expand=True)
        render_folders()

        # --- ВКЛАДКА: АУДИО И СЕТЬ (ТОЛЬКО ДЛЯ АДМИНА) ---
        if is_admin and tab_audio_net is not None:
            ctk.CTkLabel(tab_audio_net, text="Провайдер Аудио:").pack(pady=(10, 0))
            audio_provider_var = ctk.StringVar(value=self.global_settings.get("audio_provider", "OpenRouter"))

            audio_model_entry = ctk.CTkComboBox(tab_audio_net, width=450, values=["Загрузка..."])
            audio_model_entry.set("Загрузка...")

            def update_audio_models(choice):
                # Временно блокируем ввод, пока идет загрузка
                audio_model_entry.configure(state="disabled")
                audio_model_entry.set("Загрузка моделей...")

                def fetch_and_update():
                    if choice == "Groq":
                        models = ["whisper-large-v3-turbo", "whisper-large-v3", "distil-whisper-large-v3-en"]
                    else:
                        # Базовый список 100% проверенных аудиомоделей
                        models = [
                            "openai/gpt-4o-audio-preview",
                            "openai/gpt-audio-mini",
                            "google/gemini-2.5-flash",
                            "google/gemini-2.0-flash-001"
                        ]
                        try:
                            resp = requests.get("https://openrouter.ai/api/v1/models", timeout=8)
                            data = resp.json().get("data", [])
                            for m in data:
                                m_id = str(m.get("id", ""))
                                m_id_low = m_id.lower()
                                # Ищем слова-маркеры в названии модели
                                if "audio" in m_id_low or "whisper" in m_id_low or "voxtral" in m_id_low or "mimo" in m_id_low:
                                    if m_id not in models:
                                        models.append(m_id)
                        except Exception as e:
                            print(f"Ошибка загрузки моделей OpenRouter: {e}")
                    
                    audio_model_entry.winfo_toplevel().after(0, lambda: apply_models(models, choice))

                def apply_models(models, current_choice):
                    audio_model_entry.configure(values=models, state="normal")

                    # Пытаемся восстановить ранее сохраненную модель
                    saved_model = self.global_settings.get("audio_model", "")
                    saved_provider = self.global_settings.get("audio_provider", "OpenRouter")

                    if current_choice == saved_provider and saved_model in models:
                        audio_model_entry.set(saved_model)
                    else:
                        audio_model_entry.set(models[0])

                # Запускаем загрузку в фоне, чтобы интерфейс не зависал
                import threading
                threading.Thread(target=fetch_and_update, daemon=True).start()

            audio_provider_menu = ctk.CTkOptionMenu(
                tab_audio_net,
                variable=audio_provider_var,
                values=["OpenRouter", "Groq"],
                command=update_audio_models
            )
            audio_provider_menu.pack(pady=5)

            ctk.CTkLabel(tab_audio_net, text="Модель Аудио (можно вписать свою):").pack(pady=(10, 0))
            audio_model_entry.pack(pady=5)

            # Первичная инициализация списка при открытии окна
            update_audio_models(audio_provider_var.get())

            ctk.CTkLabel(tab_audio_net, text="Длина куска (мин):").pack(pady=(10, 0))
            audio_chunk_entry = ctk.CTkEntry(tab_audio_net, width=450)
            audio_chunk_entry.pack(pady=5)
            audio_chunk_entry.insert(0, str(self.global_settings.get("audio_chunk_mins", 60)))

            ctk.CTkLabel(tab_audio_net, text="Перекрытие (сек):").pack(pady=(10, 0))
            audio_overlap_entry = ctk.CTkEntry(tab_audio_net, width=450)
            audio_overlap_entry.pack(pady=5)
            audio_overlap_entry.insert(0, str(self.global_settings.get("audio_overlap_secs", 15)))

            proxy_checkbox = ctk.CTkCheckBox(tab_audio_net, text="Использовать SOCKS5 Proxy")
            proxy_checkbox.pack(pady=(14, 5), anchor="w", padx=50)
            if bool(self.current_settings.get("use_proxy", False)):
                proxy_checkbox.select()
            else:
                proxy_checkbox.deselect()

            ctk.CTkLabel(tab_audio_net, text="Proxy Host:").pack(pady=(10, 0))
            proxy_host_entry = ctk.CTkEntry(tab_audio_net, width=450)
            proxy_host_entry.pack(pady=5)
            proxy_host_entry.insert(0, str(self.current_settings.get("proxy_host", "127.0.0.1")))

            ctk.CTkLabel(tab_audio_net, text="Proxy Port:").pack(pady=(10, 0))
            proxy_port_entry = ctk.CTkEntry(tab_audio_net, width=450)
            proxy_port_entry.pack(pady=5)
            proxy_port_entry.insert(0, str(self.current_settings.get("proxy_port", "2080")))

        # --- ВКЛАДКА 4: О ПРОГРАММЕ ---
        ctk.CTkLabel(tab_about, text=APP_NAME, font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(20, 8))
        ctk.CTkLabel(tab_about, text=f"Версия: {APP_VERSION}", font=ctk.CTkFont(size=14)).pack(pady=4)
        ctk.CTkLabel(tab_about, text=f"Разработчик: {APP_DEVELOPER} | вн. тел.: {APP_PHONE}", font=ctk.CTkFont(size=14)).pack(pady=4)
        ctk.CTkLabel(
            tab_about,
            text=APP_DESCRIPTION,
            justify="center",
            wraplength=450
        ).pack(pady=(12, 10), padx=20)

        # --- СОХРАНЕНИЕ НАСТРОЕК ---
        def save():
            new_model = model_entry.get().strip()
            if is_admin:
                self.current_settings["admin_model"] = new_model
            else:
                self.current_settings["guest_model"] = new_model

            if is_admin:
                # Настройки Аудио и Прокси (СОХРАНЯЕТ ТОЛЬКО АДМИН)
                self.global_settings["audio_provider"] = audio_provider_var.get()
                self.global_settings["audio_model"] = audio_model_entry.get().strip()
                try:
                    self.global_settings["audio_chunk_mins"] = int((audio_chunk_entry.get().strip() or "60"))
                except Exception:
                    self.global_settings["audio_chunk_mins"] = 60
                try:
                    self.global_settings["audio_overlap_secs"] = int((audio_overlap_entry.get().strip() or "15"))
                except Exception:
                    self.global_settings["audio_overlap_secs"] = 15

                self.current_settings["use_proxy"] = bool(proxy_checkbox.get())
                self.current_settings["proxy_host"] = proxy_host_entry.get().strip() or "127.0.0.1"
                self.current_settings["proxy_port"] = proxy_port_entry.get().strip() or "2080"

                # 1. Обновление истории топ-10 моделей
                history = self.current_settings.get("model_history", [])
                if new_model in history:
                    history.remove(new_model)
                history.insert(0, new_model)
                self.current_settings["model_history"] = history[:10] # Храним только 10 последних
                
                # 2. Сохранение остальных системных полей
                self.global_settings["vision_model"] = vision_entry.get().strip()
                self.global_settings["secretary_model"] = secretary_entry.get().strip()
                self.global_settings["embedding_model"] = embed_entry.get().strip()
                
                # 3. Сохранение папок и исключений
                ex_text = excludes_entry.get("1.0", "end-1c")
                self.global_settings["exclude_keywords"] = [k.strip() for k in ex_text.split(",") if k.strip()]
                self.global_settings["indexed_folders"] = temp_folders.copy()

                save_global_settings(self.global_settings)

                # 4. Сохранение Vault
                new_vault = {
                    "openrouter_key": openrouter_entry.get().strip() if openrouter_entry else "",
                    "groq_key": groq_entry.get().strip() if groq_entry else "",
                    "tavily_key": tavily_entry.get().strip() if tavily_entry else "",
                    "admin_password": (admin_pwd_entry.get().strip() if admin_pwd_entry else "admin") or "admin"
                }
                save_vault_data(new_vault)

            save_local_settings(self.current_settings)
            settings_window.destroy()

        save_btn = ctk.CTkButton(settings_window, text="Сохранить", command=save, fg_color="#2E7D32", hover_color="#1B5E20")
        save_btn.pack(pady=(10, 20))

    # ==================== ОПРЕДЕЛЕНИЕ ИНСТРУМЕНТОВ ====================
    def get_tools_schema(self):
        tools = [
            {
                "type": "function",
                "function": {
                    "name": "list_available_files",
                    "description": "Умный навигатор по папкам. Выдает структурированный список всех доступных файлов. Вызывай этот инструмент, если пользователь говорит: 'поищи в папках', 'какие есть файлы', 'найди все аудиофайлы', 'есть ли у нас схемы', или ищет файл по слову в названии.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "category": {
                                "type": "string",
                                "enum": ["all", "audio", "excel", "word", "pdf", "image", "text", "diagram"],
                                "description": "Тип искомых файлов. Используй 'all', если пользователь не назвал конкретный тип."
                            },
                            "search_keyword": {
                                "type": "string",
                                "description": "Слово для поиска в названии файла (например, 'транскрибация', 'отчет'). Оставь пустым для вывода всех файлов."
                            }
                        }
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "read_local_file",
                    "description": "Прочитать текст из файла (Поддерживает .docx, .doc, .rtf, .txt, .md, .pdf, .png, .jpg, .jpeg, .xlsx, .xls, .graphml блок-схемы, а также .mp3/.wav/.m4a/.ogg с системной меткой транскрибации). Для PDF/изображений используется smart vision роутер с кэшем. ВАЖНО: Если тебе нужно узнать содержимое директории, передай сюда путь к папке (например 'SMK_Docs/Протоколы'), и инструмент вернет тебе список файлов внутри нее.",
                    "parameters": {
                        "type": "object", 
                        "properties": {
                            "filename": {"type": "string", "description": "Имя файла или путь к папке"}
                        }, 
                        "required": ["filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "transcribe_audio_file",
                    "description": "Запустить процесс текстовой расшифровки (транскрибации) аудиофайла (.mp3, .wav, .m4a). Вызывает нейросеть для распознавания голоса и создает Word-документ с протоколом. ВНИМАНИЕ: Процесс долгий и платный. Вызывать СТРОГО только после получения явного согласия пользователя!",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {"type": "string", "description": "Имя аудиофайла (например, запись_совещания.mp3)"}
                        },
                        "required": ["filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "search_smk_knowledge_base",
                    "description": "Искать стандарты, правила и факты памяти в единой базе.",
                    "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "web_search_tavily",
                    "description": "Искать информацию, новости, статьи и актуальные требования во всем интернете. Вызывать ТОЛЬКО если пользователь дал прямое согласие на поиск в интернете (Tavily).",
                    "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "search_wikipedia",
                    "description": "Искать термины, общие знания и определения в Википедии. Вызывать ТОЛЬКО если пользователь дал прямое согласие на поиск в Википедии.",
                    "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "recall_past_conversation",
                    "description": "Вспомнить старые детали ТЕКУЩЕГО диалога. Вызывай, если пользователь ссылается на то, что вы обсуждали ранее ('как мы решили ту проблему?', 'какой процесс мы проверяли?'), но ты не видишь этого в текущей истории чата.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "query": {"type": "string", "description": "О чем именно нужно вспомнить (ключевые слова)."}
                        },
                        "required": ["query"]
                    }
                }
            },

            {
                "type": "function",
                "function": {
                    "name": "generate_mermaid_diagram",
                    "description": "Создать HTML-файл диаграммы Mermaid по готовому коду. Используй для схем, диаграмм, структур, алгоритмов и mindmap. Для mindmap: используй ТОЛЬКО отступы для иерархии, без стрелок/связей, без стилей/классов и без слова root.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "title": {"type": "string", "description": "Название схемы (станет именем файла)."},
                            "mermaid_code": {"type": "string", "description": "Код Mermaid без объяснений."}
                        },
                        "required": ["title", "mermaid_code"]
                    }
                }
            },

            {
                "type": "function",
                "function": {
                    "name": "smart_excel_search",
                    "description": f"Найти конкретные строки (проблемы, несоответствия) в таблице Excel. По умолчанию ищи в файле '{self.global_settings.get('default_excel_file', '')}'.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {"type": "string", "description": f"Имя файла Excel (по умолчанию '{self.global_settings.get('default_excel_file', '')}')."},
                            "task_description": {"type": "string", "description": "Кого или что ищем (отдел, суть проблемы)."},
                            "only_open": {"type": "boolean", "description": "Установи true, если нужно найти ТОЛЬКО актуальные/открытые/нерешенные проблемы."}
                        },
                        "required": ["filename", "task_description"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "smart_excel_edit",
                    "description": "Обновить старую или создать новую строку в таблице Excel. Вызывай ТОЛЬКО с согласия пользователя.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {"type": "string", "description": "Имя файла Excel."},
                            "task_description": {"type": "string", "description": "Что нужно обновить или создать (например 'измени статус на Выполнено' или 'добавь новую')."},
                            "found_context_str": {"type": "string", "description": "Сюда передай Топ-5 строк, которые тебе вернул инструмент smart_excel_search. Если создаешь новую запись с нуля, передай '[]'."}
                        },
                        "required": ["filename", "task_description", "found_context_str"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "apply_indexed_edits",
                    "description": "МАССОВО заменяет или удаляет абзацы в Word по их номерам (индексам). ОБЯЗАТЕЛЬНО передавай ВСЕ правки в одном вызове (в виде массива edits_list).",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "filename": {"type": "string", "description": "Имя файла"},
                            "edits_list": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "target_indices": {"type": "array", "items": {"type": "integer"}, "description": "Массив индексов абзацев для изменения (например [14, 15])"},
                                        "new_text": {"type": "string", "description": "Новый текст. Если нужно только удалить, пиши 'delete'"}
                                    },
                                    "required": ["target_indices", "new_text"]
                                }
                            }
                        },
                        "required": ["filename", "edits_list"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "generate_document_from_template",
                    "description": "Создать НОВЫЙ документ на основе файла-образца (шаблона). Используй, когда просят составить план, протокол или отчет на основе старого.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "template_filename": {"type": "string", "description": "Имя файла-образца (например, План_аудита_старый.docx)"},
                            "task_description": {"type": "string", "description": "Что именно нужно изменить (процесс, даты, ФИО)"},
                            "new_filename": {"type": "string", "description": "Имя для нового файла (например, Новый_План.docx)"}
                        },
                        "required": ["template_filename", "task_description", "new_filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "generate_document_from_scratch",
                    "description": "Разработать АБСОЛЮТНО НОВЫЙ документ С НУЛЯ (например: 'разработай новую политику', 'напиши инструкцию'). Генерирует новую структуру.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "task_description": {"type": "string", "description": "Подробное описание того, что нужно написать (какие разделы, пункты)"},
                            "new_filename": {"type": "string", "description": "Имя для нового файла"},
                            "reference_filename": {"type": "string", "description": "(Опционально) Имя файла для копирования стилей и шапки"}
                        },
                        "required": ["task_description", "new_filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "generate_excel_from_scratch",
                    "description": "Создать АБСОЛЮТНО НОВУЮ таблицу Excel с нуля. Поддерживает создание МНОГОСТРАНИЧНЫХ таблиц (несколько листов). Используй, когда пользователь просит 'сделать табличку', 'создать эксель'.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "task_description": {"type": "string", "description": "Подробное описание структуры: какие листы, какие колонки и какие данные нужны в строках."},
                            "new_filename": {"type": "string", "description": "Имя для нового файла (с расширением .xlsx)"}
                        },
                        "required": ["task_description", "new_filename"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "draft_email",
                    "description": "Создать черновик электронного письма в Outlook для отправки коллегам (информирование о несоответствиях, отправка отчетов).",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "to_name": {"type": "string", "description": "Имя получателя или 'Укажите email'."},
                            "subject": {"type": "string", "description": "Тема письма."},
                            "html_body": {"type": "string", "description": "Текст письма в строгом корпоративном HTML (используй <p>, <ul>, <li>, <strong>)."}
                        },
                        "required": ["to_name", "subject", "html_body"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "draft_meeting",
                    "description": "Создать приглашение на встречу в Outlook (назначить аудит, разбор проблем).",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "to_name": {"type": "string", "description": "Имя участников."},
                            "subject": {"type": "string", "description": "Тема встречи."},
                            "body": {"type": "string", "description": "Повестка встречи ОБЫЧНЫМ ТЕКСТОМ. КАТЕГОРИЧЕСКИ БЕЗ HTML-ТЕГОВ! Используй переносы строк (\\n) и тире для списков."},
                            "duration_minutes": {"type": "integer", "description": "Длительность в минутах."}
                        },
                        "required": ["to_name", "subject", "body", "duration_minutes"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "generate_yed_diagram",
                    "description": "Создать yEd GraphML-схему (блок-схема, процесс, маршрут, IDEF-подобная структура) с узлами, группами и связями. Для иерархии используй shape='group' и вложенный массив nodes. Пример: [{\"id\":\"g1\",\"label\":\"Группа 1\",\"shape\":\"group\",\"nodes\":[{\"id\":\"n1\",\"label\":\"Шаг 1\",\"shape\":\"process\"},{\"id\":\"n2\",\"label\":\"Решение\",\"shape\":\"decision\"}]}]",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "title": {"type": "string", "description": "Название схемы (станет именем файла .graphml)."},
                            "nodes": {
                                "type": "array",
                                "description": "Массив узлов схемы.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "id": {"type": "string", "description": "Уникальный ID узла (например, n1)."},
                                        "label": {"type": "string", "description": "Подпись узла."},
                                        "shape": {
                                            "type": "string",
                                            "enum": ["start", "end", "process", "decision", "document", "database", "manual_input", "actor", "routing", "idef_node", "group"],
                                            "description": "Тип фигуры yEd."},
                                        "nodes": {
                                            "type": "array",
                                            "description": "Вложенные узлы (используется только для shape=group).",
                                            "items": {
                                                "type": "object",
                                                "properties": {
                                                    "id": {"type": "string", "description": "Уникальный ID узла (например, n1)."},
                                                    "label": {"type": "string", "description": "Подпись узла."},
                                                    "shape": {
                                                        "type": "string",
                                                        "enum": ["start", "end", "process", "decision", "document", "database", "manual_input", "actor", "routing", "idef_node", "group"],
                                                        "description": "Тип фигуры yEd."
                                                    }
                                                },
                                                "required": ["id", "label", "shape"]
                                            }
                                        }
                                    },
                                    "required": ["id", "label", "shape"]
                                }
                            },
                            "edges": {
                                "type": "array",
                                "description": "Массив связей между узлами.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "source": {"type": "string", "description": "ID узла-источника."},
                                        "target": {"type": "string", "description": "ID узла-назначения."},
                                        "label": {"type": "string", "description": "Подпись над стрелкой (опционально)."},
                                        "flow_type": {
                                            "type": "string",
                                            "enum": ["material", "information"],
                                            "description": "Тип потока: material (сплошная), information (пунктир)."
                                        }
                                    },
                                    "required": ["source", "target", "flow_type"]
                                }
                            }
                        },
                        "required": ["title", "nodes", "edges"]
                    }
                }
            }
        ]

        # ДОБАВЛЯЕМ ИНСТРУМЕНТЫ ПАМЯТИ ТОЛЬКО ДЛЯ АДМИНА
        if getattr(self, "current_role", "guest") == "admin":
            tools.extend([
                {
                    "type": "function",
                    "function": {
                        "name": "memorize_important_fact",
                        "description": "Сохранить факт в корпоративную память.",
                        "parameters": {"type": "object", "properties": {"fact": {"type": "string"}}, "required": ["fact"]}
                    }
                },
                {
                    "type": "function",
                    "function": {
                        "name": "forget_fact",
                        "description": "Удалить факт из корпоративной памяти.",
                        "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}
                    }
                }
            ])

        return tools

    def execute_tool(self, func_name, args):
        if func_name == "list_available_files": return list_available_files(args.get("category", "all"), args.get("search_keyword", ""))
        elif func_name == "read_local_file": return read_local_file(args.get("filename"))
        elif func_name == "transcribe_audio_file": return transcribe_audio_logic(args.get("filename"), self)
        elif func_name == "search_smk_knowledge_base": return search_smk_knowledge_base(args.get("query"))
        elif func_name == "web_search_tavily": return web_search_tavily(args.get("query"))
        elif func_name == "search_wikipedia": return search_wikipedia_tool(args.get("query"))
        elif func_name in ["memorize_important_fact", "forget_fact"]:
            if getattr(self, "current_role", "guest") != "admin":
                return "ОШИБКА БЕЗОПАСНОСТИ: У вас нет прав Администратора для изменения корпоративной базы знаний."

            if func_name == "memorize_important_fact":
                return memorize_important_fact(args.get("fact"))
            else:
                return forget_fact(args.get("query"))
        elif func_name == "recall_past_conversation": return recall_past_conversation(args.get("query"), self)
        elif func_name == "generate_mermaid_diagram":
            return generate_mermaid_diagram(args.get("title"), args.get("mermaid_code"), self)
        elif func_name == "generate_yed_diagram":
            return generate_yed_diagram(args.get("title"), args.get("nodes"), args.get("edges"), self)

        elif func_name == "smart_excel_search": return smart_excel_search(args.get("filename"), args.get("task_description"), args.get("only_open", False), self)
        elif func_name == "smart_excel_edit": return smart_excel_edit(args.get("filename"), args.get("task_description"), args.get("found_context_str"), self)
        elif func_name == "apply_indexed_edits": return apply_indexed_edits(args.get("filename"), args.get("edits_list"))
        elif func_name == "generate_document_from_template": return generate_document_from_template(args.get("template_filename"), args.get("task_description"), args.get("new_filename"), self)
        elif func_name == "generate_document_from_scratch": return generate_document_from_scratch(args.get("task_description"), args.get("new_filename"), args.get("reference_filename", ""), self)
        elif func_name == "generate_excel_from_scratch": return generate_excel_from_scratch(args.get("task_description"), args.get("new_filename"), self)
        elif func_name == "draft_email":
            return draft_email_tool(args.get("to_name"), args.get("subject"), args.get("html_body"))
        elif func_name == "draft_meeting":
            return draft_meeting_tool(args.get("to_name"), args.get("subject"), args.get("body"), args.get("duration_minutes", 60))
        else: return f"Ошибка: Инструмент не найден."

    # ==================== АГЕНТНЫЙ ЦИКЛ ====================
    def send_message(self):
        user_text = self.input_entry.get("1.0", "end-1c").strip()
        if not user_text: return
        self.append_to_chat(f" Вы: {user_text} ", "user_msg")
        self.append_to_chat("\n\n")
        self.input_entry.delete("1.0", "end")
        self.chat_history.append({"role": "user", "content": user_text})
        self.save_history()
        # ЭШЕЛОН 6: Генерируем название сессии ТОЛЬКО если это Админ
        if len(self.chat_history) == 1 and getattr(self, "current_role", "guest") == "admin":
            threading.Thread(target=self.generate_session_title_background, args=(user_text,), daemon=True).start()
        threading.Thread(target=self.agent_loop, daemon=True).start()

    def agent_loop(self):
        self.append_to_chat(" ИИ-Агент: \n", "agent_msg")
        
        system_prompt = (
            "Ты суперинтеллектуальный автономный агент СМК.\n"
            "ТВОЙ СТРОГИЙ АЛГОРИТМ РАБОТЫ:\n"
            "ШАГ 1. СВЕРКА: При любом запросе СНАЧАЛА вызывай 'search_smk_knowledge_base'.\n"
            "ШАГ 1.1. ПРОВЕРКА ИНТЕРНЕТА: Если в локальной базе знаний нет ответа на вопрос пользователя, ты НЕ ИМЕЕШЬ ПРАВА сразу придумывать ответ или искать его в сети. Сначала напиши пользователю: 'В нашей локальной базе СМК нет этой информации. Где мне поискать ответ: в интернете (Tavily) или в Википедии?'.\n"
            "ШАГ 1.2. Дождись ответа. Если пользователь выбрал интернет - вызови 'web_search_tavily'. Если Википедию - вызови 'search_wikipedia'. ПРИ ОТВЕТЕ ИЗ ВНЕШНЕЙ СЕТИ ОБЯЗАТЕЛЬНО УКАЗЫВАЙ ПРЯМЫЕ ВЕБ-ССЫЛКИ на источники (http...).\n"
            "ШАГ 1.3. АУДИОФАЙЛЫ: Если пользователь просит тебя проанализировать или пересказать аудиофайл, ВЫЗОВИ инструмент 'read_local_file' с именем этого аудио. Инструмент сам достанет текст из кэша. Если же в кэше пусто (инструмент вернет предупреждение), ТЫ НЕ ИМЕЕШЬ ПРАВА вызывать 'transcribe_audio_file' без разрешения. Обязательно спроси: 'Я вижу аудиофайл. Запустить расшифровку голоса в текст?'. Вызывай 'transcribe_audio_file' ТОЛЬКО после слова 'Да' от пользователя.\n"
            "ШАГ 1.4. НАВИГАЦИЯ ПО ПАПКАМ: Если пользователь задает общие вопросы вроде 'поищи в папках', 'найди все аудиофайлы', 'есть ли документы с названием X', СНАЧАЛА ОБЯЗАТЕЛЬНО вызови 'list_available_files'. Этот инструмент выдаст тебе сгруппированную структуру папок и файлов. Изучи этот список и только потом отвечай пользователю или запускай чтение/расшифровку конкретных файлов.\n"
            "ШАГ 2. ПРАВКИ В ДОКУМЕНТЕ (WORD): Если просят исправить текстовый документ, прочитай его и ТОЛЬКО после согласия вызови 'apply_indexed_edits'.\n"
            f"ШАГ 3. ТАБЛИЦЫ (EXCEL): Если просят проверить, добавить или обновить несоответствие в Excel (по умолчанию файл '{self.current_settings.get('default_excel_file', '')}'):\n"
            "   А) СНАЧАЛА ОБЯЗАТЕЛЬНО вызови 'smart_excel_search', чтобы найти контекст и старые записи.\n"
            "   Б) После получения результатов, покажи их пользователю и получи согласие.\n"
            "   В) Вызови 'smart_excel_edit' для внесения изменений.\n"
            "ШАГ 4. ГЕНЕРАЦИЯ ПО ШАБЛОНУ: Если просят создать документ ПО ОБРАЗЦУ, используй 'generate_document_from_template'.\n"
            "ШАГ 5. СОЗДАНИЕ С НУЛЯ: Если просят разработать/создать АБСОЛЮТНО НОВЫЙ документ, используй 'generate_document_from_scratch' (для текстовых документов Word) ИЛИ 'generate_excel_from_scratch' (для таблиц, планов и матриц в Excel).\n"
            "ШАГ 6. ВИЗУАЛИЗАЦИЯ И СХЕМЫ: Доступно 2 инструмента — 'generate_yed_diagram' (формат yEd GraphML) и 'generate_mermaid_diagram' (формат Mermaid HTML). Если пользователь не указал формат явно, сначала ОБЯЗАТЕЛЬНО спроси, какой формат нужен (yEd GraphML или Mermaid), дождись ответа и только затем вызывай соответствующий инструмент.\n"
            "ШАГ 7. КОММУНИКАЦИЯ (OUTLOOK): Если после аудита, записи в журнал или генерации отчета тебе нужно оповестить коллег или назначить разбор полетов, ВЫЗОВИ 'draft_email' (для писем с красивым HTML) или 'draft_meeting' (для встреч строгим плоским текстом). Если email адресата не указан явно, пиши просто ФИО.\n"
            "ШАГ 8. БЕСКОНЕЧНАЯ ПАМЯТЬ: Ты помнишь только последние 20 сообщений. Если пользователь ссылается на старые детали диалога, которых нет в текущей истории, ВЫЗОВИ инструмент 'recall_past_conversation'. НЕ используй его для поиска стандартов (для этого есть 'search_smk_knowledge_base').\n"
            "ШАГ 9. КЛИКАБЕЛЬНЫЕ ССЫЛКИ НА ФАЙЛЫ: Если ты упоминаешь документ СМК, нашел его через поиск или отредактировал, ОБЯЗАТЕЛЬНО выводи пользователю ссылку в строгом формате: [Из файла: Имя_файла.ext]. НИКОГДА не пытайся писать сетевые пути (\\\\Server\\...) или обычные markdown-ссылки, используй ТОЛЬКО формат в квадратных скобках!\n"
        )
        
        messages_for_llm = [{"role": "system", "content": system_prompt}] + self.chat_history
        
        for step in range(10):
            try:
                start_index = self.chat_textbox.index("end-1c")
                if getattr(self, "current_role", "guest") == "admin":
                    current_model = self.current_settings.get("admin_model", "openai/gpt-4o-mini")
                else:
                    current_model = self.current_settings.get("guest_model", "stepfun/step-3.5-flash:free")
                response = get_llm_client().chat.completions.create(
                    model=current_model,
                    messages=messages_for_llm,
                    tools=self.get_tools_schema(),
                    stream=True
                )

                content_parts = []
                tool_calls_acc = {}

                for chunk in response:
                    if not chunk.choices:
                        continue
                    delta = chunk.choices[0].delta

                    if delta.content is not None:
                        content_parts.append(delta.content)
                        self.append_to_chat(delta.content)

                    if delta.tool_calls:
                        for tc in delta.tool_calls:
                            tc_index = tc.index if tc.index is not None else 0
                            if tc_index not in tool_calls_acc:
                                tool_calls_acc[tc_index] = {
                                    "id": tc.id or f"tool_call_{tc_index}",
                                    "type": tc.type or "function",
                                    "function": {"name": "", "arguments": ""}
                                }

                            current_tc = tool_calls_acc[tc_index]

                            if tc.id:
                                current_tc["id"] = tc.id
                            if tc.type:
                                current_tc["type"] = tc.type
                            if tc.function:
                                if tc.function.name:
                                    current_tc["function"]["name"] += tc.function.name
                                if tc.function.arguments:
                                    current_tc["function"]["arguments"] += tc.function.arguments

                final_text = "".join(content_parts)
                merged_tool_calls = [tool_calls_acc[idx] for idx in sorted(tool_calls_acc.keys())]

                assistant_message = {"role": "assistant"}
                if final_text:
                    assistant_message["content"] = final_text
                if merged_tool_calls:
                    assistant_message["tool_calls"] = merged_tool_calls
                messages_for_llm.append(assistant_message)

                if not merged_tool_calls:
                    self.append_to_chat("\n\n")
                    self.apply_markdown(start_index)
                    self.chat_history.append({"role": "assistant", "content": final_text})
                    self.save_history()

                    # --- НОВОЕ: Логика вытеснения (Скользящее окно 20 сообщений = 10 пар) ---
                    if len(self.chat_history) > 20:
                        old_user = self.chat_history.pop(0)
                        old_assist = self.chat_history.pop(0)

                        # Сохраняем в векторную базу ТОЛЬКО для Админа
                        if getattr(self, "current_role", "guest") == "admin":
                            try:
                                archive_text = f"Пользователь: {old_user.get('content', '')}\nАссистент: {old_assist.get('content', '')}"
                                client = chromadb.PersistentClient(path=get_db_path())
                                collection = client.get_or_create_collection(name="temp_chat_memory", embedding_function=get_cloud_ef())
                                collection.add(
                                    documents=[archive_text],
                                    metadatas=[{"session_id": self.current_session_id}],
                                    ids=[str(uuid.uuid4())]
                                )
                            except Exception as e:
                                print(f"Ошибка архивации чата: {e}")

                    self.save_current_session()
                    break

                for tool_call in merged_tool_calls:
                    func_name = tool_call.get("function", {}).get("name", "")
                    args_raw = tool_call.get("function", {}).get("arguments", "{}")

                    try:
                        args = json.loads(args_raw) if args_raw else {}
                    except Exception:
                        args = {}

                    # Выводим аккуратный лог действия с отступом, БЕЗ дублирования бейджа
                    self.after(0, self.append_to_chat, f"  ⚙️ [Действие: {func_name}]...\n")
                    tool_result = self.execute_tool(func_name, args)
                    messages_for_llm.append({
                        "role": "tool",
                        "tool_call_id": tool_call.get("id", ""),
                        "name": func_name,
                        "content": str(tool_result)
                    })
                     
            except Exception as e:
                self.append_to_chat(f"\n[Критическая ошибка Агента: {str(e)}]\n\n")
                self.save_current_session()
                break
        else:
            # ЭТОТ БЛОК СРАБОТАЕТ ТОЛЬКО ЕСЛИ АГЕНТ ИСЧЕРПАЛ 10 ШАГОВ
            warning_msg = "⚠️ ИИ-Агент: Достигнут лимит размышлений (10 шагов). Задача слишком объемная, либо я не могу найти нужные данные. Пожалуйста, уточните запрос."
            self.append_to_chat(f"\n{warning_msg}\n\n")
            self.chat_history.append({"role": "assistant", "content": warning_msg})
            self.save_history()
            self.save_current_session()
        
        # Выбираем последние 4 сообщения для контекста
        recent_msgs = self.chat_history[-4:]
        threading.Thread(target=self.run_background_secretary, args=(recent_msgs,), daemon=True).start()

if __name__ == '__main__':
    app = App()
    app.mainloop()
