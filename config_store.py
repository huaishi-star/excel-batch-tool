"""规则/映射档案存储：保存与载入重命名规则、映射表。

只保存规则本身，不保存文件列表和输出文件夹：
- rename: {名称: {"coords": [...], "separator": "..."}}
- fill:   {名称: {"mapping": [["A1", "B2"], ...]}}

配置文件位置：
- Windows：软件所在目录（便携，随 exe 拷贝带走）；不可写时回退到 %APPDATA%\\Lsy的Excel-Tool\\
- macOS：~/Library/Application Support/Lsy的Excel-Tool/（.app 内部通常不可写，必须放用户目录）
- Linux：$XDG_CONFIG_HOME/Lsy的Excel-Tool/（默认 ~/.config/Lsy的Excel-Tool/）
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

CONFIG_NAME = 'Lsy的Excel-Tool_profiles.json'
KINDS = ('rename', 'fill', 'delete')

# 测试可覆盖的基准目录；为 None 时使用默认位置
_BASE: Path | None = None


def set_base_dir(path) -> None:
    """仅测试用：把配置读写目录固定到指定位置；传 None 恢复默认。"""
    global _BASE
    _BASE = None if path is None else Path(path)


def _default_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _appdata_dir() -> Path:
    """各平台的用户配置目录。"""
    if sys.platform == 'darwin':
        return Path.home() / 'Library' / 'Application Support' / 'Lsy的Excel-Tool'
    if sys.platform.startswith('linux'):
        base = os.environ.get('XDG_CONFIG_HOME', str(Path.home() / '.config'))
        return Path(base) / 'Lsy的Excel-Tool'
    return Path(os.environ.get('APPDATA', str(Path.home()))) / 'Lsy的Excel-Tool'


def config_paths() -> list[Path]:
    if _BASE is not None:
        return [_BASE / CONFIG_NAME]
    default = _default_dir() / CONFIG_NAME
    appdata = _appdata_dir() / CONFIG_NAME
    if sys.platform == 'darwin':
        # macOS 的 .app 内部可能只读：优先用户数据目录，同时兼容脚本旁的旧配置
        return [appdata, default]
    return [default, appdata]


def _load() -> dict:
    """读取全部档案；文件缺失或损坏时返回空结构。"""
    empty = {'rename': {}, 'fill': {}, 'delete': {}}
    for p in config_paths():
        if not p.exists():
            continue
        try:
            data = json.loads(p.read_text(encoding='utf-8'))
        except Exception:
            continue
        if isinstance(data, dict):
            return {
                'rename': data.get('rename') or {},
                'fill': data.get('fill') or {},
                'delete': data.get('delete') or {},
            }
    return empty


def _save(data: dict) -> None:
    text = json.dumps(data, ensure_ascii=False, indent=2)
    last_err: Exception | None = None
    for p in config_paths():
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(text, encoding='utf-8')
            return
        except OSError as e:
            last_err = e
    raise last_err


def list_profiles(kind: str) -> list[str]:
    """返回指定类型（rename/fill）所有已保存记录的名称。"""
    return sorted(_load().get(kind, {}).keys())


def get_profile(kind: str, name: str) -> dict | None:
    """取指定记录；不存在返回 None。"""
    return _load().get(kind, {}).get(name)


def save_profile(kind: str, name: str, payload: dict) -> None:
    """保存（或覆盖）一条记录。"""
    data = _load()
    data.setdefault(kind, {})[name] = payload
    _save(data)


def delete_profile(kind: str, name: str) -> None:
    """删除一条记录。"""
    data = _load()
    data.get(kind, {}).pop(name, None)
    _save(data)


def clear_config() -> None:
    """删除所有位置的配置文件（测试/重置用）。"""
    for p in config_paths():
        try:
            if p.exists():
                p.unlink()
        except OSError:
            pass
