"""Persistent JSON document projects (MIT adapted workflow; safe project identifiers)."""
from __future__ import annotations
import copy
import json
import os
from pathlib import Path
import re
import tempfile
import uuid
from datetime import datetime, timezone


def workspace() -> Path:
    value = os.environ.get('WPS_CLI_WORKSPACE', '').strip()
    if value:
        base = Path(value).expanduser()
    elif os.name == 'nt':
        base = Path(os.environ.get('LOCALAPPDATA') or Path.home() / 'AppData' / 'Local') / 'cli-anything-wps' / 'projects'
    else:
        # Linux is not a supported execution target; this path enables offline model tests.
        base = Path(tempfile.gettempdir()) / 'cli-anything-wps-model-test'
    base.mkdir(parents=True, exist_ok=True)
    return base.resolve()


def _path(pid: str) -> Path:
    if not isinstance(pid, str) or not re.fullmatch(r'[0-9a-f]{32}', pid):
        raise ValueError('无效的 project_id，请使用 project_create 返回的 32 位 ID')
    return workspace() / (pid + '.json')


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def create(kind: str, name: str, width: float = 960, height: float = 540) -> dict:
    if kind not in ('writer', 'calc', 'impress'):
        raise ValueError('doc_type 只能是 writer、calc、impress')
    pid = uuid.uuid4().hex
    obj = {'project_id': pid, 'type': kind, 'name': (name or '未命名')[:200],
           'created_at': _now(), 'updated_at': _now(), 'version': 1,
           'undo_stack': [], 'redo_stack': []}
    if kind == 'writer': obj['content'] = []
    if kind == 'calc': obj['sheets'] = [{'name': 'Sheet1', 'cells': {}, 'merges': []}]
    if kind == 'impress':
        obj['slides'] = []
        obj['slide_width'] = width
        obj['slide_height'] = height
    save(obj)
    return summary(obj)


def summary(obj: dict) -> dict:
    d = {k: obj[k] for k in ('project_id', 'type', 'name', 'created_at', 'updated_at') if k in obj}
    if obj['type'] == 'writer': d['content_count'] = len(obj.get('content', []))
    if obj['type'] == 'calc': d['sheet_count'] = len(obj.get('sheets', []))
    if obj['type'] == 'impress': d['slide_count'] = len(obj.get('slides', []))
    return d


def load(pid: str) -> dict:
    p = _path(pid)
    if not p.is_file(): raise FileNotFoundError(f'项目不存在: {pid}')
    with p.open('r', encoding='utf-8') as f: return json.load(f)


def save(obj: dict) -> None:
    target = _path(obj['project_id'])
    obj['updated_at'] = _now()
    fd, temp = tempfile.mkstemp(prefix='.wps-', suffix='.tmp', dir=target.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp, target)
    finally:
        if os.path.exists(temp): os.unlink(temp)


def snapshot(obj: dict) -> None:
    content_key = {'writer': 'content', 'calc': 'sheets', 'impress': 'slides'}[obj['type']]
    obj.setdefault('undo_stack', []).append(copy.deepcopy(obj[content_key]))
    obj['undo_stack'] = obj['undo_stack'][-30:]
    obj['redo_stack'] = []


def undo_redo(pid: str, undo: bool = True) -> dict:
    obj = load(pid)
    source, target = ('undo_stack', 'redo_stack') if undo else ('redo_stack', 'undo_stack')
    if not obj.get(source): raise ValueError('没有可撤销/重做的操作')
    k = {'writer': 'content', 'calc': 'sheets', 'impress': 'slides'}[obj['type']]
    obj.setdefault(target, []).append(copy.deepcopy(obj[k]))
    obj[k] = obj[source].pop()
    save(obj)
    return summary(obj)


def all_projects() -> list[dict]:
    result = []
    for p in sorted(workspace().glob('*.json'), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            with p.open('r', encoding='utf-8') as f: result.append(summary(json.load(f)))
        except (ValueError, KeyError, OSError): continue
    return result[:200]


def remove(pid: str) -> dict:
    p = _path(pid)
    if not p.exists(): raise FileNotFoundError(f'项目不存在: {pid}')
    p.unlink()
    return {'deleted': pid}
