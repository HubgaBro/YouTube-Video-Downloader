import time
APP_START_TIME = time.perf_counter()
APP_VERSION = "2026.4.24.0"
SETTINGS_FILE_NAME = "settings.json"

import sys
import importlib
import json
import os
import re
import urllib.parse
import urllib.request
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, 
    QProgressBar, QMessageBox, QFileDialog, QListWidget,
    QDialog, QDialogButtonBox, QListWidgetItem, QSplashScreen,
    QGroupBox, QFormLayout
)
from PySide6.QtCore import QThread, Signal, QUrl, Qt, QTimer
from PySide6.QtGui import QDesktopServices, QPixmap, QPainter


_yt_dlp_module = None
_yt_dlp_import_error = None


def get_yt_dlp_module():
    global _yt_dlp_module, _yt_dlp_import_error

    if _yt_dlp_module is not None:
        return _yt_dlp_module

    if _yt_dlp_import_error is not None:
        raise RuntimeError(f"Failed to import yt_dlp: {_yt_dlp_import_error}") from _yt_dlp_import_error

    try:
        _yt_dlp_module = importlib.import_module('yt_dlp')
    except Exception as exc:
        _yt_dlp_import_error = exc
        raise RuntimeError(f"Failed to import yt_dlp: {exc}") from exc

    return _yt_dlp_module


def format_duration(seconds):
    try:
        total_seconds = int(float(seconds))
    except (TypeError, ValueError):
        return "-"

    if total_seconds <= 0:
        return "-"

    hours, remainder = divmod(total_seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours > 0:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    return f"{minutes}:{secs:02d}"


def format_resolution(info):
    resolution = info.get('resolution')
    if resolution and resolution != 'audio only':
        return str(resolution)

    width = info.get('width')
    height = info.get('height')
    if width and height:
        return f"{width}x{height}"

    best_dimensions = None
    for fmt in info.get('formats') or []:
        if fmt.get('vcodec') == 'none':
            continue
        fmt_width = fmt.get('width')
        fmt_height = fmt.get('height')
        if not fmt_width or not fmt_height:
            continue
        if best_dimensions is None or fmt_height > best_dimensions[1]:
            best_dimensions = (fmt_width, fmt_height)

    if best_dimensions:
        return f"{best_dimensions[0]}x{best_dimensions[1]}"
    return "-"


def format_size_text(size_bytes):
    try:
        size_value = float(size_bytes)
    except (TypeError, ValueError):
        return "-"

    if size_value <= 0:
        return "-"

    units = ['B', 'KB', 'MB', 'GB', 'TB']
    unit_index = 0
    while size_value >= 1024 and unit_index < len(units) - 1:
        size_value /= 1024.0
        unit_index += 1
    return f"{size_value:.1f} {units[unit_index]}"


def get_browser_cookie_option(browser_key):
    if not browser_key or browser_key == 'none':
        return None
    return (browser_key,)


def apply_cookie_settings(ydl_opts, cookie_settings=None):
    if not cookie_settings:
        return

    browser_option = cookie_settings.get('browser_option')
    if browser_option:
        ydl_opts['cookiesfrombrowser'] = browser_option

    cookie_header = (cookie_settings.get('cookie_header') or '').strip()
    if cookie_header:
        headers = dict(ydl_opts.get('http_headers') or {})
        headers['Cookie'] = cookie_header
        ydl_opts['http_headers'] = headers


def _estimate_best_muxed_filesize(info):
    requested_downloads = info.get('requested_downloads') or []
    if requested_downloads:
        total_bytes = 0
        has_size = False
        for item in requested_downloads:
            part_size = item.get('filesize') or item.get('filesize_approx') or 0
            if part_size:
                has_size = True
                total_bytes += part_size
        if has_size:
            return total_bytes

    best_video = None
    best_audio = None
    best_progressive = None
    for fmt in info.get('formats') or []:
        fmt_size = fmt.get('filesize') or fmt.get('filesize_approx')
        if not fmt_size:
            continue

        vcodec = fmt.get('vcodec')
        acodec = fmt.get('acodec')
        if vcodec != 'none' and acodec != 'none':
            if best_progressive is None or fmt_size > best_progressive:
                best_progressive = fmt_size
            continue

        if vcodec != 'none':
            if best_video is None or fmt.get('height', 0) > best_video[0]:
                best_video = (fmt.get('height', 0), fmt_size)
        elif acodec != 'none':
            abr = fmt.get('abr') or fmt.get('tbr') or 0
            if best_audio is None or abr > best_audio[0]:
                best_audio = (abr, fmt_size)

    if best_video and best_audio:
        return best_video[1] + best_audio[1]
    if best_progressive:
        return best_progressive
    return info.get('filesize') or info.get('filesize_approx')


def _estimate_audio_transcode_size(duration_seconds, bitrate_kbps):
    try:
        duration_value = float(duration_seconds)
    except (TypeError, ValueError):
        return None

    if duration_value <= 0 or bitrate_kbps <= 0:
        return None
    return int(duration_value * bitrate_kbps * 1000 / 8)


def estimate_size_map(info):
    duration_seconds = info.get('duration')
    best_audio_size = None
    best_audio_bitrate = None
    for fmt in info.get('formats') or []:
        if fmt.get('acodec') == 'none':
            continue
        audio_size = fmt.get('filesize') or fmt.get('filesize_approx')
        audio_bitrate = fmt.get('abr') or fmt.get('tbr') or 0
        if audio_size and (best_audio_size is None or audio_bitrate > (best_audio_bitrate or 0)):
            best_audio_size = audio_size
            best_audio_bitrate = audio_bitrate

    return {
        'mp4_hd': _estimate_best_muxed_filesize(info),
        'mp3_high': _estimate_audio_transcode_size(duration_seconds, 320) or best_audio_size,
        'mp3_low': _estimate_audio_transcode_size(duration_seconds, 128) or best_audio_size,
        'wav_lossless': _estimate_audio_transcode_size(duration_seconds, 1411),
        'ogg_lossy': _estimate_audio_transcode_size(duration_seconds, 192) or best_audio_size,
    }


def resolve_entry_url(entry):
    if not isinstance(entry, dict):
        return None

    for key in ('webpage_url', 'original_url', 'url'):
        value = entry.get(key)
        if isinstance(value, str) and value.startswith('http'):
            return value

    video_id = entry.get('id')
    extractor = str(entry.get('extractor') or entry.get('ie_key') or '').lower()
    if video_id and ('youtube' in extractor or not extractor):
        return f"https://www.youtube.com/watch?v={video_id}"
    return None


def fetch_thumbnail_bytes(thumbnail_url):
    if not thumbnail_url:
        return None

    request = urllib.request.Request(
        thumbnail_url,
        headers={'User-Agent': 'Mozilla/5.0'}
    )
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            return response.read()
    except Exception:
        return None


def is_supported_youtube_url(url):
    if not url:
        return False

    try:
        parsed = urllib.parse.urlparse(url.strip())
    except Exception:
        return False

    if parsed.scheme not in ('http', 'https'):
        return False

    hostname = (parsed.hostname or '').lower()
    if hostname.startswith('www.'):
        hostname = hostname[4:]
    if hostname.startswith('m.'):
        hostname = hostname[2:]

    if hostname == 'youtu.be':
        return bool(parsed.path.strip('/'))

    if hostname == 'youtube.com' or hostname.endswith('.youtube.com'):
        path = parsed.path or ''
        if path.startswith('/watch'):
            return 'v=' in parsed.query
        if path.startswith('/playlist'):
            return 'list=' in parsed.query
        if path.startswith('/shorts/') or path.startswith('/live/') or path.startswith('/embed/'):
            return True

    return False


def extract_youtube_url(text):
    if not text:
        return None

    raw_text = text.strip()
    if is_supported_youtube_url(raw_text):
        return raw_text

    for match in re.findall(r'https?://[^\s<>"\']+', raw_text):
        candidate = match.rstrip(').,;!?]')
        if is_supported_youtube_url(candidate):
            return candidate

    return None


def extract_media_preview(url, cookie_settings=None):
    yt_dlp = get_yt_dlp_module()
    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'skip_download': True,
        'extract_flat': 'in_playlist',
    }

    apply_cookie_settings(ydl_opts, cookie_settings)

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)

    is_playlist = info.get('_type') == 'playlist'
    entries = []
    for index, entry in enumerate(info.get('entries') or [], start=1):
        if not entry:
            continue

        entry_url = resolve_entry_url(entry)
        if not entry_url:
            continue

        entries.append({
            'index': index,
            'title': entry.get('title') or f"Video {index}",
            'url': entry_url,
            'duration_text': format_duration(entry.get('duration')),
        })

    return {
        'source_url': url,
        'is_playlist': is_playlist,
        'title': info.get('title') or info.get('playlist_title') or info.get('fulltitle') or '-',
        'uploader': info.get('uploader') or info.get('channel') or info.get('playlist_uploader') or info.get('uploader_id') or '-',
        'duration_text': '-' if is_playlist else format_duration(info.get('duration')),
        'resolution_text': '-' if is_playlist else format_resolution(info),
        'size_map': {} if is_playlist else estimate_size_map(info),
        'thumbnail_bytes': fetch_thumbnail_bytes(info.get('thumbnail')),
        'entry_count': info.get('playlist_count') or len(entries) or (1 if not is_playlist else 0),
        'entries': entries,
    }


class MediaInfoThread(QThread):
    ready = Signal(object)
    error = Signal(str)

    def __init__(self, url, cookie_settings=None):
        super().__init__()
        self.url = url
        self.cookie_settings = cookie_settings

    def run(self):
        try:
            self.ready.emit(extract_media_preview(self.url, self.cookie_settings))
        except Exception as e:
            self.error.emit(str(e))


class YtDlpWarmupThread(QThread):
    def run(self):
        try:
            get_yt_dlp_module()
        except Exception:
            pass


class PlaylistSelectionDialog(QDialog):
    def __init__(self, playlist_title, entries, is_english=False, parent=None):
        super().__init__(parent)
        self.entries = entries
        self.is_english = is_english

        self.setWindowTitle("Select Playlist Items" if is_english else "選擇播放清單項目")
        self.resize(560, 420)

        layout = QVBoxLayout(self)

        title = playlist_title or ('Playlist' if is_english else '播放清單')
        summary = (
            f"Select the videos to download from: {title}"
            if is_english else
            f"請勾選要下載的影片：{title}"
        )
        summary_label = QLabel(summary)
        summary_label.setWordWrap(True)
        layout.addWidget(summary_label)

        controls = QHBoxLayout()
        select_all_btn = QPushButton("Select All" if is_english else "全選")
        select_all_btn.clicked.connect(lambda: self._set_all(Qt.CheckState.Checked))
        controls.addWidget(select_all_btn)

        clear_all_btn = QPushButton("Clear All" if is_english else "全部取消")
        clear_all_btn.clicked.connect(lambda: self._set_all(Qt.CheckState.Unchecked))
        controls.addWidget(clear_all_btn)
        controls.addStretch()
        layout.addLayout(controls)

        self.list_widget = QListWidget()
        for entry in self.entries:
            item_text = f"{entry['index']}. {entry['title']}"
            if entry.get('duration_text') and entry['duration_text'] != '-':
                item_text = f"{item_text} ({entry['duration_text']})"
            item = QListWidgetItem(item_text)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _set_all(self, state):
        for index in range(self.list_widget.count()):
            self.list_widget.item(index).setCheckState(state)

    def selected_entries(self):
        selected = []
        for index, entry in enumerate(self.entries):
            if self.list_widget.item(index).checkState() == Qt.CheckState.Checked:
                selected.append(entry)
        return selected


class SettingsDialog(QDialog):
    def __init__(
        self,
        cookie_options,
        language_index,
        output_dir,
        cookie_source,
        manual_cookie,
        app_version,
        ffmpeg_path,
        parent=None,
    ):
        super().__init__(parent)
        self.cookie_options = cookie_options
        self.app_version = app_version
        self.ffmpeg_path = ffmpeg_path
        self.initial_cookie_source = cookie_source or 'none'

        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setMinimumWidth(460)
        self.resize(500, 430)

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(18, 18, 18, 18)

        self.panel = QWidget()
        self.panel.setObjectName("settingsPanel")
        self.panel.setStyleSheet(
            "QWidget#settingsPanel {"
            " background-color: #f7f7f7;"
            " border: 1px solid #b8b8b8;"
            " border-radius: 12px;"
            " color: #202020;"
            "}"
            "QGroupBox {"
            " background-color: #ffffff;"
            " border: 1px solid #d7d7d7;"
            " border-radius: 8px;"
            " color: #202020;"
            " font-weight: 600;"
            " margin-top: 12px;"
            " padding-top: 10px;"
            "}"
            "QGroupBox::title {"
            " subcontrol-origin: margin;"
            " left: 12px;"
            " padding: 0 4px;"
            " color: #202020;"
            " background-color: #f7f7f7;"
            "}"
            "QLabel {"
            " color: #202020;"
            "}"
            "QLineEdit, QComboBox {"
            " color: #202020;"
            " background-color: #ffffff;"
            " border: 1px solid #c8c8c8;"
            " border-radius: 6px;"
            " padding: 6px 8px;"
            "}"
            "QPushButton {"
            " color: #202020;"
            " background-color: #ffffff;"
            " border: 1px solid #c8c8c8;"
            " border-radius: 6px;"
            " padding: 6px 12px;"
            "}"
            "QPushButton:hover {"
            " background-color: #f0f0f0;"
            "}"
            "QPushButton:pressed {"
            " background-color: #e7e7e7;"
            "}"
        )
        outer_layout.addWidget(self.panel)

        layout = QVBoxLayout(self.panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        self.title_label = QLabel()
        layout.addWidget(self.title_label)

        self.summary_label = QLabel()
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.general_group = QGroupBox()
        general_layout = QFormLayout(self.general_group)

        self.language_row_label = QLabel()
        self.language_combo = QComboBox()
        self.language_combo.addItems(["繁體中文", "English"])
        self.language_combo.setCurrentIndex(language_index)
        general_layout.addRow(self.language_row_label, self.language_combo)

        self.output_dir_row_label = QLabel()
        output_dir_row = QWidget()
        output_dir_layout = QHBoxLayout(output_dir_row)
        output_dir_layout.setContentsMargins(0, 0, 0, 0)
        self.output_dir_input = QLineEdit(output_dir)
        self.output_dir_browse_button = QPushButton()
        output_dir_layout.addWidget(self.output_dir_input, 1)
        output_dir_layout.addWidget(self.output_dir_browse_button)
        general_layout.addRow(self.output_dir_row_label, output_dir_row)
        layout.addWidget(self.general_group)

        self.download_group = QGroupBox()
        download_layout = QFormLayout(self.download_group)

        self.cookie_source_row_label = QLabel()
        self.cookie_combo = QComboBox()
        download_layout.addRow(self.cookie_source_row_label, self.cookie_combo)

        self.manual_cookie_row_label = QLabel()
        self.manual_cookie_input = QLineEdit(manual_cookie or "")
        download_layout.addRow(self.manual_cookie_row_label, self.manual_cookie_input)

        self.ffmpeg_status_row_label = QLabel()
        self.ffmpeg_status_value = QLabel()
        self.ffmpeg_status_value.setWordWrap(True)
        download_layout.addRow(self.ffmpeg_status_row_label, self.ffmpeg_status_value)
        layout.addWidget(self.download_group)

        self.about_group = QGroupBox()
        about_layout = QVBoxLayout(self.about_group)

        self.version_button = QPushButton()
        about_layout.addWidget(self.version_button)

        links_layout = QHBoxLayout()
        self.website_button = QPushButton()
        self.github_button = QPushButton("GitHub")
        links_layout.addWidget(self.website_button)
        links_layout.addWidget(self.github_button)
        about_layout.addLayout(links_layout)
        layout.addWidget(self.about_group)

        layout.addStretch()

        self.button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self.output_dir_browse_button.clicked.connect(self._browse_output_dir)
        self.language_combo.currentIndexChanged.connect(self._on_language_changed)
        self.cookie_combo.currentIndexChanged.connect(self._update_manual_cookie_visibility)
        self.version_button.clicked.connect(self._show_version)

        self._refresh_ui_texts()
        self._populate_cookie_combo(self.initial_cookie_source)
        self._update_manual_cookie_visibility()

    def _tr(self, zh_text, en_text):
        return en_text if self.language_combo.currentIndex() == 1 else zh_text

    def _on_language_changed(self):
        self._refresh_ui_texts()

    def _populate_cookie_combo(self, selected_key=None):
        current_key = selected_key or self.cookie_combo.currentData() or self.initial_cookie_source
        self.cookie_combo.blockSignals(True)
        self.cookie_combo.clear()

        for key, zh_text, en_text in self.cookie_options:
            self.cookie_combo.addItem(self._tr(zh_text, en_text), key)

        selected_index = self.cookie_combo.findData(current_key)
        if selected_index < 0:
            selected_index = 0
        self.cookie_combo.setCurrentIndex(selected_index)
        self.cookie_combo.blockSignals(False)
        self._update_manual_cookie_visibility()

    def _refresh_ffmpeg_status(self):
        if self.ffmpeg_path:
            self.ffmpeg_status_value.setText(self._tr("已偵測到", "Detected"))
            self.ffmpeg_status_value.setStyleSheet("color: #1f9d3a; font-weight: 600;")
            self.ffmpeg_status_value.setToolTip(self.ffmpeg_path)
        else:
            self.ffmpeg_status_value.setText(self._tr("未偵測到", "Not Detected"))
            self.ffmpeg_status_value.setStyleSheet("color: #d12f2f; font-weight: 600;")
            self.ffmpeg_status_value.setToolTip("")

    def _refresh_ui_texts(self):
        current_cookie_source = self.cookie_combo.currentData() or self.initial_cookie_source

        self.setWindowTitle(self._tr("設定面板", "Settings Panel"))
        self.title_label.setText(self._tr("設定面板", "Settings Panel"))
        self.title_label.setStyleSheet("font-size: 18px; font-weight: 700; color: #202020;")
        self.summary_label.setText(
            self._tr(
                "在這裡調整應用程式偏好設定，並查看下載環境與版本資訊。",
                "Adjust application preferences here and review download environment and version details."
            )
        )
        self.summary_label.setStyleSheet("color: #505050;")

        self.general_group.setTitle(self._tr("一般", "General"))
        self.language_row_label.setText(self._tr("語言", "Language"))
        self.output_dir_row_label.setText(self._tr("輸出資料夾", "Output Folder"))
        self.output_dir_input.setPlaceholderText(self._tr("選擇預設輸出資料夾", "Choose a default output folder"))
        self.output_dir_browse_button.setText(self._tr("瀏覽", "Browse"))

        self.download_group.setTitle(self._tr("下載", "Download"))
        self.cookie_source_row_label.setText(self._tr("瀏覽器 Cookie", "Browser Cookies"))
        self.manual_cookie_row_label.setText(self._tr("手動 Cookie", "Manual Cookie"))
        self.manual_cookie_input.setPlaceholderText(
            self._tr("貼上 Cookie（例如：SID=...; HSID=...）", "Paste Cookie (example: SID=...; HSID=...)")
        )
        self.ffmpeg_status_row_label.setText(self._tr("FFmpeg 狀態", "FFmpeg Status"))
        self._refresh_ffmpeg_status()

        self.about_group.setTitle(self._tr("關於", "About"))
        self.version_button.setText(self._tr("版本", "Version"))
        self.website_button.setText(self._tr("官方網站", "Official Website"))

        save_button = self.button_box.button(QDialogButtonBox.Save)
        if save_button is not None:
            save_button.setText(self._tr("儲存", "Save"))

        cancel_button = self.button_box.button(QDialogButtonBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText(self._tr("取消", "Cancel"))

        self._populate_cookie_combo(current_cookie_source)

    def _update_manual_cookie_visibility(self):
        show_manual_input = self.cookie_combo.currentData() == 'manual'
        self.manual_cookie_row_label.setVisible(show_manual_input)
        self.manual_cookie_input.setVisible(show_manual_input)

    def _browse_output_dir(self):
        initial_dir = self.output_dir_input.text().strip() or str(Path.home())
        folder = QFileDialog.getExistingDirectory(
            self,
            self._tr("選擇輸出資料夾", "Select Output Directory"),
            initial_dir,
        )
        if folder:
            self.output_dir_input.setText(folder)

    def _show_version(self):
        QMessageBox.information(
            self,
            self._tr("版本資訊", "Version Information"),
            self._tr(f"目前版本：{self.app_version}", f"Current version: {self.app_version}"),
        )

    def selected_settings(self):
        return {
            'language_index': self.language_combo.currentIndex(),
            'output_dir': self.output_dir_input.text().strip(),
            'cookie_source': self.cookie_combo.currentData() or 'none',
            'manual_cookie': self.manual_cookie_input.text(),
        }

class DownloadThread(QThread):
    progress = Signal(int)
    info = Signal(object)
    completed = Signal(str)
    error = Signal(str)

    def __init__(self, url, format_selector, output_path, ffmpeg_location=None, custom_filename=None, cookie_settings=None):
        super().__init__()
        self.url = url
        self.format_selector = format_selector
        self.output_path = output_path
        self.ffmpeg_location = ffmpeg_location
        self.custom_filename = self._sanitize_filename(custom_filename)
        self.cookie_settings = cookie_settings
        self.last_filename = None
        self._is_cancelled = False

    def _sanitize_filename(self, filename):
        if not filename:
            return None

        sanitized = re.sub(r'[<>:"/\\|?*]', '', filename).strip().rstrip('.')
        if not sanitized:
            return None
        return sanitized

    def _build_output_template(self):
        filename_part = '%(title)s'
        if self.custom_filename:
            filename_part = self.custom_filename.replace('%', '%%')
        return str(Path(self.output_path) / f'{filename_part}.%(ext)s')

    def _extract_percent(self, progress_data):
        downloaded = progress_data.get('downloaded_bytes') or 0
        total = progress_data.get('total_bytes') or progress_data.get('total_bytes_estimate') or 0

        if total > 0:
            return min((downloaded / total) * 100, 100.0)

        percent_str = str(progress_data.get('_percent_str', '') or '')
        match = re.search(r'(\d+(?:\.\d+)?)\s*%', percent_str)
        if match:
            return min(float(match.group(1)), 100.0)

        return 0.0

    def _get_output_extension(self):
        extension_map = {
            'mp4_hd': 'mp4',
            'mp3_high': 'mp3',
            'mp3_low': 'mp3',
            'wav_lossless': 'wav',
            'ogg_lossy': 'ogg',
        }
        return extension_map.get(self.format_selector)

    def _build_final_output_path(self, filename):
        if not filename:
            return None

        output_extension = self._get_output_extension()
        output_path = Path(filename)
        if output_extension:
            return str(output_path.with_suffix(f".{output_extension}"))
        return str(output_path)

    def run(self):
        yt_dlp = get_yt_dlp_module()

        def progress_hook(d):
            if getattr(self, '_is_cancelled', False):
                raise Exception('Cancelled by user')
            if d.get('status') == 'downloading':
                try:
                    p = self._extract_percent(d)
                    eta = d.get('eta')
                    downloaded = d.get('downloaded_bytes') or 0
                    total = d.get('total_bytes') or d.get('total_bytes_estimate') or 0

                    self.progress.emit(int(p))
                    self.info.emit({
                        'stage': 'downloading',
                        'percent': int(p),
                        'eta': eta,
                        'downloaded': downloaded,
                        'total': total,
                        'filename': self.last_filename,
                    })
                except Exception as e:
                    self.error.emit(f"Error: progress update failed: {str(e)}")

            if d.get('status') == 'finished':
                self.info.emit({'stage': 'download_finished'})

        common_opts = {
            'outtmpl': self._build_output_template(),
            'progress_hooks': [progress_hook],
            'quiet': True,
            'no_warnings': True,
            'retries': 3,
            'fragment_retries': 3,
            'concurrent-fragments': 8,
            'embed-metadata': True,
        }

        if self.ffmpeg_location:
            common_opts['ffmpeg_location'] = self.ffmpeg_location
        apply_cookie_settings(common_opts, self.cookie_settings)

        if self.format_selector == 'mp4_hd':
            ydl_opts = {
                **common_opts,
                'format': 'bestvideo+bestaudio/best',
                'merge_output_format': 'mp4',
                'postprocessor_args': {'FFmpegMerger': ['-movflags', '+faststart']},
            }
        elif self.format_selector == 'mp3_high':
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '0',
                }],
            }
        elif self.format_selector == 'mp3_low':
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '5',
                }],
            }
        elif self.format_selector == 'wav_lossless':
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'wav',
                }],
            }
        elif self.format_selector == 'ogg_lossy':
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'vorbis',
                    'preferredquality': '192',
                }],
            }
        else:
            ydl_opts = common_opts

        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(self.url, download=True)

            if isinstance(info, dict) and info.get('_type') == 'playlist':
                entries = info.get('entries') or []
                info = entries[0] if entries else info

            if info:
                prepared_filename = ydl.prepare_filename(info)
                self.last_filename = self._build_final_output_path(prepared_filename)

            self.info.emit({'stage': 'done', 'filename': self.last_filename})
            self.completed.emit(self.last_filename or "")
        except Exception as e:
            msg = str(e)
            if 'Cancelled' in msg or 'cancel' in msg.lower():
                self.error.emit(f"Cancelled: {msg}")
            else:
                self.error.emit(f"Error: {msg}")

    def request_cancel(self):
        self._is_cancelled = True

class MainWindow(QMainWindow):
    FORMAT_OPTIONS = (
        ("mp4_hd", "MP4 | 高解析度", "MP4 | High Resolution"),
        ("mp3_high", "MP3 | 高位元率", "MP3 | High Bitrate"),
        ("mp3_low", "MP3 | 低位元率", "MP3 | Low Bitrate"),
        ("wav_lossless", "WAV | 無損", "WAV | Lossless"),
        ("ogg_lossy", "OGG | 有損", "OGG | Lossy"),
    )
    COOKIE_BROWSER_OPTIONS = (
        ("none", "不使用瀏覽器 Cookie", "No Browser Cookies"),
        ("manual", "手動貼上 Cookie", "Paste Cookie Manually"),
        ("chrome", "Google Chrome", "Google Chrome"),
        ("edge", "Microsoft Edge", "Microsoft Edge"),
        ("firefox", "Mozilla Firefox", "Mozilla Firefox"),
        ("brave", "Brave", "Brave"),
        ("opera", "Opera", "Opera"),
        ("vivaldi", "Vivaldi", "Vivaldi"),
        ("chromium", "Chromium", "Chromium"),
    )

    def __init__(self):
        super().__init__()
        self.setWindowTitle("YouTube Video Downloader")
        self.setMinimumSize(760, 700)
        self.resize(900, 760)

        self.app_version = APP_VERSION
        self.settings_ready = False
        self.startup_complete = False
        self.status = QLabel("程式正在啟動...")
        self.ffmpeg_path = self.find_ffmpeg()
        self.preview_data = None
        self.preview_url = None
        self.active_preview_url = None
        self.pending_preview_url = None
        self.preview_loading = False
        self.preview_thread = None
        self.is_downloading = False
        self.can_open_folder = False
        self.can_open_file = False
        self.last_downloaded_file = None
        self.pending_queue_start = False
        self.yt_dlp_warmup_thread = None

        self.auto_preview_timer = QTimer(self)
        self.auto_preview_timer.setSingleShot(True)
        self.auto_preview_timer.timeout.connect(self._trigger_auto_preview)
        self.clipboard = QApplication.clipboard()
        self.clipboard.dataChanged.connect(self.on_clipboard_changed)

        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        layout.setSpacing(8)
        layout.setContentsMargins(25, 25, 25, 25)

        top_bar = QHBoxLayout()
        self.ffmpeg_status_label = QLabel()
        top_bar.addWidget(self.ffmpeg_status_label)
        top_bar.addStretch()
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["繁體中文", "English"])
        top_bar.addWidget(self.lang_combo)
        layout.addLayout(top_bar)

        self.label = QLabel("Paste YouTube URL:")
        layout.addWidget(self.label)

        url_layout = QHBoxLayout()
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://www.youtube.com/watch?v=...")
        self.url_input.textChanged.connect(self.on_url_changed)
        url_layout.addWidget(self.url_input)
        layout.addLayout(url_layout)

        filename_layout = QHBoxLayout()
        self.filename_label = QLabel("File Name:")
        filename_layout.addWidget(self.filename_label)

        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("Leave blank to use the default title")
        filename_layout.addWidget(self.filename_input)
        layout.addLayout(filename_layout)

        self.mode_label = QLabel("Choose Download Format:")
        layout.addWidget(self.mode_label)

        self.format_combo = QComboBox()
        self.format_combo.currentIndexChanged.connect(self._refresh_preview_panel)
        layout.addWidget(self.format_combo)

        cookie_layout = QHBoxLayout()
        self.cookie_label = QLabel("Browser Cookies:")
        cookie_layout.addWidget(self.cookie_label)

        self.cookie_combo = QComboBox()
        self.cookie_combo.currentIndexChanged.connect(self.on_cookie_source_changed)
        cookie_layout.addWidget(self.cookie_combo)
        layout.addLayout(cookie_layout)

        self.cookie_input_label = QLabel("Cookie Header:")
        layout.addWidget(self.cookie_input_label)

        self.cookie_input = QLineEdit()
        self.cookie_input.setPlaceholderText("貼上 Cookie（例如：SID=...; HSID=...）")
        self.cookie_input.textChanged.connect(self.on_manual_cookie_changed)
        layout.addWidget(self.cookie_input)

        self.preview_section_label = QLabel("Media Preview:")
        layout.addWidget(self.preview_section_label)

        preview_layout = QHBoxLayout()
        self.thumbnail_label = QLabel("No Preview")
        self.thumbnail_label.setAlignment(Qt.AlignCenter)
        self.thumbnail_label.setMinimumSize(176, 99)
        self.thumbnail_label.setMaximumSize(352, 198)
        self.thumbnail_label.setStyleSheet("border: 1px solid #808080; background-color: #202020; color: #f0f0f0;")
        preview_layout.addWidget(self.thumbnail_label)

        preview_text_layout = QVBoxLayout()
        self.preview_type_line = QLabel()
        self.preview_type_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_type_line)

        self.preview_title_line = QLabel()
        self.preview_title_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_title_line)

        self.preview_author_line = QLabel()
        self.preview_author_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_author_line)

        self.preview_duration_line = QLabel()
        self.preview_duration_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_duration_line)

        self.preview_resolution_line = QLabel()
        self.preview_resolution_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_resolution_line)

        self.preview_size_line = QLabel()
        self.preview_size_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_size_line)

        self.preview_note_line = QLabel("")
        self.preview_note_line.setWordWrap(True)
        preview_text_layout.addWidget(self.preview_note_line)
        preview_text_layout.addStretch()

        preview_layout.addLayout(preview_text_layout, 1)
        layout.addLayout(preview_layout, 1)

        queue_ctrl = QHBoxLayout()
        self.btn_add_queue = QPushButton("加入佇列")
        self.btn_add_queue.clicked.connect(self.add_to_queue)
        queue_ctrl.addWidget(self.btn_add_queue)

        self.btn_start_queue = QPushButton("開始佇列")
        self.btn_start_queue.clicked.connect(self.start_queue)
        queue_ctrl.addWidget(self.btn_start_queue)

        self.btn_clear_queue = QPushButton("清空佇列")
        self.btn_clear_queue.clicked.connect(self.clear_queue)
        queue_ctrl.addWidget(self.btn_clear_queue)

        self.btn_cancel = QPushButton("取消下載")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_current_download)
        queue_ctrl.addWidget(self.btn_cancel)

        queue_ctrl.addStretch()
        layout.addLayout(queue_ctrl)

        self.queue_list = QListWidget()
        self.queue_list.setMaximumHeight(180)
        layout.addWidget(self.queue_list)

        self.btn_download = QPushButton("Download")
        self.btn_download.clicked.connect(self.start_download)
        layout.addWidget(self.btn_download)

        progress_layout = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setMaximum(100)
        self.progress.setTextVisible(False)
        progress_layout.addWidget(self.progress)
        
        self.progress_label = QLabel("0%")
        self.progress_label.setMinimumWidth(40)
        progress_layout.addWidget(self.progress_label)
        layout.addLayout(progress_layout)

        layout.addWidget(self.status)
        info_layout = QHBoxLayout()
        self.stage_label = QLabel("")
        self.stage_label.setMinimumWidth(180)
        info_layout.addWidget(self.stage_label)

        self.file_label = QLabel("")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        info_layout.addWidget(self.file_label)

        self.size_eta_label = QLabel("")
        self.size_eta_label.setMinimumWidth(160)
        self.size_eta_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        info_layout.addWidget(self.size_eta_label)

        layout.addLayout(info_layout)

        open_layout = QHBoxLayout()
        self.btn_open_folder = QPushButton("開啟資料夾")
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self.open_folder)
        self.btn_open_folder.setToolTip("在檔案管理器中開啟輸出資料夾")
        self.btn_open_folder.setVisible(False)
        open_layout.addWidget(self.btn_open_folder)

        self.btn_open_file = QPushButton("開啟檔案")
        self.btn_open_file.setEnabled(False)
        self.btn_open_file.clicked.connect(self.open_file)
        self.btn_open_file.setToolTip("使用預設程式開啟下載的檔案")
        self.btn_open_file.setVisible(False)
        open_layout.addWidget(self.btn_open_file)

        open_layout.addStretch()
        layout.addLayout(open_layout)
        
        bottom_layout = QHBoxLayout()
        self.btn_website = QPushButton("Official Website")
        self.btn_website.clicked.connect(self.open_website)
        bottom_layout.addWidget(self.btn_website)
        self.btn_github = QPushButton("GitHub")
        self.btn_github.clicked.connect(self.open_github)
        bottom_layout.addWidget(self.btn_github)
        self.btn_settings = QPushButton("Settings")
        self.btn_settings.clicked.connect(self.open_settings_dialog)
        bottom_layout.addWidget(self.btn_settings)
        bottom_layout.addStretch()
        self.startup_time_label = QLabel("")
        self.startup_time_label.setMinimumWidth(70)
        self.startup_time_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.startup_time_label.setStyleSheet("color: #606060;")
        bottom_layout.addWidget(self.startup_time_label)
        layout.addLayout(bottom_layout)

        loaded_settings = self._load_settings()
        self.output_dir = loaded_settings['output_dir']
        self.queue = []
        self.thread = None
        self.lang_combo.currentIndexChanged.connect(self.change_language)
        self._apply_settings(loaded_settings)
        self._update_ffmpeg_status_label()
        self._update_cookie_input_visibility()
        self.status.setText(self._startup_status_text())
        self._refresh_preview_panel()
        self._update_button_states()
        self.settings_ready = True
        QTimer.singleShot(0, self._finish_startup)
        QTimer.singleShot(0, self._warm_up_yt_dlp)

    def _tr(self, zh_text, en_text):
        return en_text if self.lang_combo.currentIndex() == 1 else zh_text

    def _settings_file_path(self):
        return Path(__file__).with_name(SETTINGS_FILE_NAME)

    def _default_settings(self):
        return {
            'language_index': 1,
            'output_dir': '下載',
            'cookie_source': 'none',
            'manual_cookie': '',
        }

    def _normalize_settings(self, raw_settings=None):
        settings = self._default_settings()
        if not isinstance(raw_settings, dict):
            return settings

        language_index = raw_settings.get('language_index')
        if language_index in (0, 1):
            settings['language_index'] = language_index

        output_dir = raw_settings.get('output_dir')
        if isinstance(output_dir, str):
            output_dir = output_dir.strip()
            if output_dir:
                settings['output_dir'] = output_dir

        valid_cookie_sources = {key for key, _, _ in self.COOKIE_BROWSER_OPTIONS}
        cookie_source = raw_settings.get('cookie_source')
        if cookie_source in valid_cookie_sources:
            settings['cookie_source'] = cookie_source

        manual_cookie = raw_settings.get('manual_cookie')
        if isinstance(manual_cookie, str):
            settings['manual_cookie'] = manual_cookie

        return settings

    def _load_settings(self):
        settings_path = self._settings_file_path()
        if not settings_path.exists():
            return self._default_settings()

        try:
            raw_settings = json.loads(settings_path.read_text(encoding='utf-8'))
        except Exception:
            return self._default_settings()

        return self._normalize_settings(raw_settings)

    def _current_settings(self):
        return self._normalize_settings({
            'language_index': self.lang_combo.currentIndex(),
            'output_dir': self.output_dir,
            'cookie_source': self.cookie_combo.currentData(),
            'manual_cookie': self.cookie_input.text(),
        })

    def _save_settings(self, settings=None, show_error=False):
        settings_to_save = self._normalize_settings(settings or self._current_settings())
        try:
            self._settings_file_path().write_text(
                json.dumps(settings_to_save, ensure_ascii=False, indent=2),
                encoding='utf-8',
            )
        except Exception as exc:
            if show_error:
                QMessageBox.warning(
                    self,
                    self._tr("錯誤", "Error"),
                    self._tr(f"無法儲存設定：{exc}", f"Unable to save settings: {exc}"),
                )
            return False

        return True

    def _persist_settings_if_ready(self):
        if not self.settings_ready:
            return
        self._save_settings()

    def _apply_cookie_preferences(self, cookie_source, manual_cookie):
        cookie_index = self.cookie_combo.findData(cookie_source)
        if cookie_index < 0:
            cookie_index = 0

        self.cookie_combo.blockSignals(True)
        self.cookie_combo.setCurrentIndex(cookie_index)
        self.cookie_combo.blockSignals(False)

        self.cookie_input.blockSignals(True)
        self.cookie_input.setText(manual_cookie)
        self.cookie_input.blockSignals(False)
        self._update_cookie_input_visibility()

    def _apply_settings(self, settings):
        normalized_settings = self._normalize_settings(settings)
        self.output_dir = normalized_settings['output_dir']

        language_index = normalized_settings['language_index']
        if self.lang_combo.currentIndex() != language_index:
            self.lang_combo.setCurrentIndex(language_index)
        else:
            self.change_language(language_index)

        self._apply_cookie_preferences(
            normalized_settings['cookie_source'],
            normalized_settings['manual_cookie'],
        )

    def _populate_format_combo(self):
        selected_key = self.format_combo.currentData()
        self.format_combo.blockSignals(True)
        self.format_combo.clear()

        for key, zh_text, en_text in self.FORMAT_OPTIONS:
            self.format_combo.addItem(self._tr(zh_text, en_text), key)

        if selected_key is not None:
            selected_index = self.format_combo.findData(selected_key)
            if selected_index >= 0:
                self.format_combo.setCurrentIndex(selected_index)

        self.format_combo.blockSignals(False)

    def _populate_cookie_combo(self):
        selected_key = self.cookie_combo.currentData()
        self.cookie_combo.blockSignals(True)
        self.cookie_combo.clear()

        for key, zh_text, en_text in self.COOKIE_BROWSER_OPTIONS:
            self.cookie_combo.addItem(self._tr(zh_text, en_text), key)

        if selected_key is None:
            selected_key = 'none'

        selected_index = self.cookie_combo.findData(selected_key)
        if selected_index < 0:
            selected_index = 0
        self.cookie_combo.setCurrentIndex(selected_index)
        self.cookie_combo.blockSignals(False)

    def _selected_cookie_settings(self):
        source_key = self.cookie_combo.currentData()
        if source_key == 'manual':
            return {
                'cookie_header': self.cookie_input.text().strip()
            }

        return {
            'browser_option': get_browser_cookie_option(source_key)
        }

    def _format_label(self, format_selector):
        for key, zh_text, en_text in self.FORMAT_OPTIONS:
            if key == format_selector:
                return self._tr(zh_text, en_text)
        return str(format_selector or "")

    def _selected_cookie_label(self):
        current_key = self.cookie_combo.currentData()
        for key, zh_text, en_text in self.COOKIE_BROWSER_OPTIONS:
            if key == current_key:
                return self._tr(zh_text, en_text)
        return self._tr("不使用瀏覽器 Cookie", "No Browser Cookies")

    def _update_cookie_input_visibility(self):
        show_manual_input = self.cookie_combo.currentData() == 'manual'
        self.cookie_input_label.setVisible(show_manual_input)
        self.cookie_input.setVisible(show_manual_input)

    def _idle_status_text(self):
        if self.ffmpeg_path:
            return self._tr("準備下載...", "Download Ready. . .")
        return self._tr(
            "準備下載... (未找到 ffmpeg，轉檔功能可能失效)",
            "Download Ready. . . (ffmpeg not found, conversion features may not work)"
        )

    def _startup_status_text(self):
        return self._tr("程式正在啟動...", "Application is starting...")

    def _update_ffmpeg_status_label(self):
        if self.ffmpeg_path:
            self.ffmpeg_status_label.setText(self._tr("FFmpeg：已偵測到", "FFmpeg: Detected"))
            self.ffmpeg_status_label.setStyleSheet("color: #1f9d3a; font-weight: 600;")
        else:
            self.ffmpeg_status_label.setText(self._tr("FFmpeg：未偵測到", "FFmpeg: Not Detected"))
            self.ffmpeg_status_label.setStyleSheet("color: #d12f2f; font-weight: 600;")

    def set_startup_time(self, elapsed_seconds):
        if elapsed_seconds is None:
            self.startup_time_label.setText("")
            return

        self.startup_time_label.setText(f"({max(elapsed_seconds, 0.0):.1f} s)")

    def _finish_startup(self):
        self.startup_complete = True
        if not self.is_downloading and not self.preview_loading and not self.preview_data and not self.url_input.text().strip():
            self.status.setText(self._idle_status_text())

    def _warm_up_yt_dlp(self):
        if self.yt_dlp_warmup_thread is not None:
            return

        self.yt_dlp_warmup_thread = YtDlpWarmupThread(self)
        self.yt_dlp_warmup_thread.finished.connect(self._on_yt_dlp_warmup_finished)
        self.yt_dlp_warmup_thread.finished.connect(self.yt_dlp_warmup_thread.deleteLater)
        self.yt_dlp_warmup_thread.start()

    def _on_yt_dlp_warmup_finished(self):
        self.yt_dlp_warmup_thread = None

    def _output_dir_dialog_title(self):
        return self._tr("選擇輸出資料夾", "Select Output Directory")

    def on_clipboard_changed(self):
        youtube_url = extract_youtube_url(self.clipboard.text())
        if not youtube_url or youtube_url == self.url_input.text().strip():
            return

        self.url_input.setText(youtube_url)

        if self.is_downloading:
            self.status.setText(
                self._tr(
                    "已從剪貼簿貼上連結，下載完成後將自動預覽",
                    "Link pasted from clipboard. Preview will start after downloads finish"
                )
            )
            return

        self.auto_preview_timer.stop()
        self.preview_media_info(youtube_url)

    def on_url_changed(self, text):
        normalized_url = text.strip()
        self.auto_preview_timer.stop()

        if normalized_url == (self.preview_url or "") and self.preview_data:
            return

        self.preview_data = None
        self.preview_url = None
        self.pending_preview_url = normalized_url or None
        self._refresh_preview_panel()

        if not normalized_url:
            if not self.is_downloading and not self.preview_loading:
                self.status.setText(self._idle_status_text())
            return

        if self.is_downloading or self.preview_loading:
            return

        self.auto_preview_timer.start(700)

    def _set_preview_loading(self, loading):
        self.preview_loading = loading
        self._update_button_states()

    def _update_button_states(self):
        if self.preview_loading:
            for button in (
                self.btn_add_queue,
                self.btn_start_queue,
                self.btn_clear_queue,
                self.btn_cancel,
                self.btn_download,
                self.btn_open_folder,
                self.btn_open_file,
                self.btn_website,
                self.btn_github,
                self.btn_settings,
            ):
                button.setEnabled(False)
            return

        self.btn_download.setEnabled(not self.is_downloading)
        self.btn_add_queue.setEnabled(not self.is_downloading)
        self.btn_start_queue.setEnabled(not self.is_downloading)
        self.btn_clear_queue.setEnabled(not self.is_downloading)
        self.btn_cancel.setEnabled(self.is_downloading)
        self.btn_open_folder.setEnabled(self.can_open_folder)
        self.btn_open_file.setEnabled(self.can_open_file)
        self.btn_website.setEnabled(True)
        self.btn_github.setEnabled(True)
        self.btn_settings.setEnabled(True)

    def _trigger_auto_preview(self):
        url = (self.pending_preview_url or self.url_input.text()).strip()
        if not url or self.is_downloading or self.preview_loading or url == self.preview_url:
            return
        self._start_preview_request(url)

    def _refresh_preview_panel(self):
        preview = self.preview_data or {}

        thumbnail_bytes = preview.get('thumbnail_bytes')
        if thumbnail_bytes:
            pixmap = QPixmap()
            if pixmap.loadFromData(thumbnail_bytes):
                scaled = pixmap.scaled(
                    self.thumbnail_label.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                self.thumbnail_label.setPixmap(scaled)
                self.thumbnail_label.setText("")
            else:
                self.thumbnail_label.setPixmap(QPixmap())
                self.thumbnail_label.setText(self._tr("無封面", "No Thumbnail"))
        else:
            self.thumbnail_label.setPixmap(QPixmap())
            self.thumbnail_label.setText(
                self._tr("無封面", "No Thumbnail") if preview else self._tr("尚未預覽", "No Preview")
            )

        if not preview:
            self.preview_type_line.setText(f"{self._tr('類型', 'Type')}: -")
            self.preview_title_line.setText(f"{self._tr('標題', 'Title')}: -")
            self.preview_author_line.setText(f"{self._tr('作者', 'Uploader')}: -")
            self.preview_duration_line.setText(f"{self._tr('時長', 'Duration')}: -")
            self.preview_resolution_line.setText(f"{self._tr('解析度', 'Resolution')}: -")
            self.preview_size_line.setText(f"{self._tr('預計大小', 'Estimated Size')}: -")
            self.preview_note_line.setText("")
            return

        is_playlist = preview.get('is_playlist', False)
        if is_playlist:
            type_value = self._tr(
                f"播放清單 ({preview.get('entry_count', 0)} 部)",
                f"Playlist ({preview.get('entry_count', 0)} items)"
            )
            duration_value = self._tr(
                f"{preview.get('entry_count', 0)} 部影片",
                f"{preview.get('entry_count', 0)} items"
            )
        else:
            type_value = self._tr("單一影片", "Single Video")
            duration_value = preview.get('duration_text') or '-'

        self.preview_type_line.setText(f"{self._tr('類型', 'Type')}: {type_value}")
        self.preview_title_line.setText(f"{self._tr('標題', 'Title')}: {preview.get('title') or '-'}")
        self.preview_author_line.setText(f"{self._tr('作者', 'Uploader')}: {preview.get('uploader') or '-'}")
        self.preview_duration_line.setText(f"{self._tr('時長', 'Duration')}: {duration_value}")
        self.preview_resolution_line.setText(
            f"{self._tr('解析度', 'Resolution')}: {preview.get('resolution_text') or '-'}"
        )

        size_map = preview.get('size_map') or {}
        selected_format = self.format_combo.currentData()
        estimated_size_text = format_size_text(size_map.get(selected_format))
        self.preview_size_line.setText(
            f"{self._tr('預計大小', 'Estimated Size')}: {estimated_size_text}"
        )
        self.preview_note_line.setText(
            self._tr(
                "偵測到播放清單，下載或加入佇列時會先讓你勾選影片。",
                "Playlist detected. You can choose entries before downloading or queueing."
            ) if is_playlist else self._tr(
                f"目前 Cookie 來源：{self._selected_cookie_label()}",
                f"Current cookie source: {self._selected_cookie_label()}"
            )
        )

    def preview_media_info(self, url=None):
        url = (url or self.url_input.text()).strip()
        if url:
            self._start_preview_request(url)

    def _start_preview_request(self, url):
        if not url:
            return

        if self.preview_loading:
            self.pending_preview_url = url
            return

        if self.preview_data and self.preview_url == url:
            return

        self.status.setText(self._tr("正在取得影片資訊...", "Loading media info..."))
        self.active_preview_url = url
        self.pending_preview_url = url
        self._set_preview_loading(True)

        self.preview_thread = MediaInfoThread(url, self._selected_cookie_settings())
        self.preview_thread.ready.connect(self.on_preview_ready)
        self.preview_thread.error.connect(self.on_preview_error)
        self.preview_thread.finished.connect(self.on_preview_finished)
        self.preview_thread.start()

    def on_preview_ready(self, preview):
        current_url = self.url_input.text().strip()
        if current_url != preview.get('source_url'):
            return

        self.preview_data = preview
        self.preview_url = preview.get('source_url')
        self._refresh_preview_panel()
        if preview.get('is_playlist'):
            self.status.setText(self._tr("已偵測到播放清單", "Playlist detected"))
        else:
            self.status.setText(self._tr("影片資訊已更新", "Media info loaded"))

    def on_preview_error(self, message):
        if self.active_preview_url != self.url_input.text().strip():
            return
        self.preview_url = self.active_preview_url
        self.status.setText(self._tr("取得影片資訊失敗", "Failed to load media info"))
        QMessageBox.critical(self, self._tr("錯誤", "Error"), message)

    def on_preview_finished(self):
        self.preview_thread = None
        self.active_preview_url = None
        self._set_preview_loading(False)

        current_url = self.url_input.text().strip()
        if current_url and current_url != (self.preview_url or "") and not self.is_downloading:
            self.pending_preview_url = current_url
            self.auto_preview_timer.start(200)

    def _ensure_preview_data(self, url):
        if self.preview_data and self.preview_url == url:
            return self.preview_data

        self.status.setText(self._tr("正在分析連結...", "Analyzing URL..."))
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            preview = extract_media_preview(url, self._selected_cookie_settings())
        except Exception as e:
            self.status.setText(self._tr("取得影片資訊失敗", "Failed to load media info"))
            QMessageBox.critical(self, self._tr("錯誤", "Error"), str(e))
            return None
        finally:
            QApplication.restoreOverrideCursor()

        self.preview_data = preview
        self.preview_url = url
        self._refresh_preview_panel()
        return preview

    def _get_queue_display_text(self, queue_item):
        label = queue_item.get('display_text') or queue_item.get('custom_filename') or queue_item['url']
        return f"{label} | {queue_item['format_selector']}"

    def _add_queue_items(self, items, prepend=False):
        if prepend:
            self.queue = list(items) + self.queue
            for index, item in enumerate(items):
                self.queue_list.insertItem(index, self._get_queue_display_text(item))
            return

        self.queue.extend(items)
        for item in items:
            self.queue_list.addItem(self._get_queue_display_text(item))

    def _select_playlist_entries(self, preview):
        entries = preview.get('entries') or []
        if not entries:
            QMessageBox.warning(
                self,
                self._tr("錯誤", "Error"),
                self._tr("無法讀取播放清單項目", "Unable to read playlist entries")
            )
            return None

        dialog = PlaylistSelectionDialog(
            preview.get('title'),
            entries,
            self.lang_combo.currentIndex() == 1,
            self
        )
        if dialog.exec() != QDialog.Accepted:
            return None

        selected_entries = dialog.selected_entries()
        if not selected_entries:
            QMessageBox.information(
                self,
                self._tr("提示", "Info"),
                self._tr("請至少勾選一部影片", "Please select at least one item")
            )
            return []
        return selected_entries

    def _build_download_items(self, url, format_selector, custom_filename=None):
        preview = self._ensure_preview_data(url)
        if not preview:
            return None

        if not preview.get('is_playlist'):
            return [{
                'url': url,
                'format_selector': format_selector,
                'custom_filename': custom_filename,
                'display_text': custom_filename or preview.get('title') or url,
            }]

        selected_entries = self._select_playlist_entries(preview)
        if selected_entries is None:
            return None
        if not selected_entries:
            return []

        if custom_filename and len(selected_entries) > 1:
            QMessageBox.information(
                self,
                self._tr("提示", "Info"),
                self._tr(
                    "播放清單多筆下載時會忽略自訂檔名，避免檔案互相覆蓋。",
                    "Custom filename is ignored when downloading multiple playlist items to avoid overwriting files."
                )
            )
            custom_filename = None

        items = []
        for entry in selected_entries:
            items.append({
                'url': entry['url'],
                'format_selector': format_selector,
                'custom_filename': custom_filename,
                'display_text': entry.get('title') or entry['url'],
            })
        return items

    def find_ffmpeg(self):
        current_dir = Path(__file__).parent.absolute()
        
        ffmpeg_exe = current_dir / "ffmpeg.exe"
        if ffmpeg_exe.exists():
            return str(ffmpeg_exe)
        
        ffmpeg_bin = current_dir / "ffmpeg"
        if ffmpeg_bin.exists():
            return str(ffmpeg_bin)
        
        return None

    def start_download(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, self._tr("錯誤", "Error"), self._tr("請輸入 YouTube 連結", "Please enter a YouTube URL"))
            return

        selected = self.format_combo.currentData()
        custom_filename = self.filename_input.text().strip() or None
        download_items = self._build_download_items(url, selected, custom_filename)
        if download_items is None or not download_items:
            return

        folder = QFileDialog.getExistingDirectory(self, self._output_dir_dialog_title(), self.output_dir)
        if not folder:
            return
        self.output_dir = folder
        self._persist_settings_if_ready()

        first_item = download_items[0]
        remaining_items = download_items[1:]
        if remaining_items:
            self._add_queue_items(remaining_items, prepend=True)
        self._start_download_task(
            first_item['url'],
            first_item['format_selector'],
            first_item['custom_filename']
        )

    def _reset_download_ui(self):
        self.progress.setValue(0)
        self.progress_label.setText("0%")
        self.file_label.setText("")
        self.size_eta_label.setText("")
        self.stage_label.setText("")
        self.can_open_folder = False
        self.can_open_file = False
        self._update_button_states()

    def _set_download_controls(self, downloading):
        self.is_downloading = downloading
        self._update_button_states()

    def _start_download_task(self, url, selected, custom_filename=None):
        self._reset_download_ui()
        self._set_download_controls(True)
        self.pending_queue_start = False
        self.last_downloaded_file = None
        selected_label = self._format_label(selected)
        self.status.setText(f"{self._tr('下載中', 'Downloading')}... ({selected_label})")

        self.thread = DownloadThread(
            url,
            selected,
            self.output_dir,
            self.ffmpeg_path,
            custom_filename,
            self._selected_cookie_settings(),
        )
        self.thread.progress.connect(self.update_progress, Qt.QueuedConnection)
        self.thread.info.connect(self.update_info, Qt.QueuedConnection)
        self.thread.completed.connect(self.on_download_completed, Qt.QueuedConnection)
        self.thread.error.connect(self.on_error, Qt.QueuedConnection)
        self.thread.finished.connect(self.on_thread_finished, Qt.QueuedConnection)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_download_completed(self, filename):
        is_english = self.lang_combo.currentIndex() == 1
        done_message = "Download complete!" if is_english else "下載完成"

        self.progress.setValue(100)
        self.progress_label.setText("100%")
        self.status.setText(done_message)
        self.last_downloaded_file = filename or None
        if self.last_downloaded_file:
            self.file_label.setText(os.path.basename(self.last_downloaded_file))
            self.can_open_folder = True
            self.can_open_file = True
            self._update_button_states()
        self.pending_queue_start = bool(self.queue)

    def on_thread_finished(self):
        finished_thread = self.sender()
        if finished_thread is self.thread:
            self.thread = None

        self._set_download_controls(False)

        if self.pending_queue_start and self.queue:
            self.pending_queue_start = False
            QTimer.singleShot(0, self.start_next_in_queue)
        else:
            self.pending_queue_start = False
            current_url = self.url_input.text().strip()
            if current_url and current_url != (self.preview_url or "") and not self.preview_loading:
                self.pending_preview_url = current_url
                self.auto_preview_timer.start(200)

    def update_progress(self, value):
        self.progress.setValue(value)
        self.progress_label.setText(f"{value}%")

    def update_info(self, info):
        stage = info.get('stage')
        if stage == 'downloading':
            eta = info.get('eta')
            total = info.get('total', 0)
            downloaded = info.get('downloaded', 0)
            percent = info.get('percent', 0)
            self.stage_label.setText(self._tr("下載中", "Downloading"))
            self.progress.setValue(percent)
            self.progress_label.setText(f"{percent}%")
            size_text = f"{self._human_size(downloaded)}/{self._human_size(total)}"
            eta_text = self._format_eta(eta)
            if eta_text:
                self.size_eta_label.setText(f"{size_text} {self._tr('預估剩餘', 'ETA')}: {eta_text}")
            else:
                self.size_eta_label.setText(size_text)
        elif stage == 'download_finished':
            self.stage_label.setText(self._tr("下載完成，後處理中", "Download finished, post-processing"))
        elif stage == 'done':
            self.stage_label.setText(self._tr("完成", "Done"))
            fn = info.get('filename')
            if fn:
                self.last_downloaded_file = fn
                self.file_label.setText(os.path.basename(fn))
                self.can_open_folder = True
                self.can_open_file = True
                self._update_button_states()

    def add_to_queue(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, self._tr("錯誤", "Error"), self._tr("請輸入 YouTube 連結", "Please enter a YouTube URL"))
            return

        queue_items = self._build_download_items(
            url,
            self.format_combo.currentData(),
            self.filename_input.text().strip() or None,
        )
        if queue_items is None or not queue_items:
            return

        self._add_queue_items(queue_items)
        self.status.setText(
            self._tr(
                f"已加入佇列 {len(queue_items)} 項",
                f"Added {len(queue_items)} item(s) to queue"
            )
        )

    def start_queue(self):
        if not self.queue:
            QMessageBox.information(self, self._tr("提示", "Info"), self._tr("佇列為空", "Queue is empty"))
            return
        if not self.output_dir or not os.path.isdir(self.output_dir):
            folder = QFileDialog.getExistingDirectory(self, self._output_dir_dialog_title(), self.output_dir)
            if not folder:
                return
            self.output_dir = folder
            self._persist_settings_if_ready()
        if not self.is_downloading:
            self.start_next_in_queue()

    def clear_queue(self):
        self.queue.clear()
        self.queue_list.clear()

    def start_next_in_queue(self):
        if not self.queue:
            return
        next_item = self.queue.pop(0)
        if self.queue_list.count() > 0:
            self.queue_list.takeItem(0)
        next_url = next_item['url']
        selected = next_item['format_selector']
        custom_filename = next_item['custom_filename']
        self._start_download_task(next_url, selected, custom_filename)

    def cancel_current_download(self):
        try:
            if hasattr(self, 'thread') and self.thread is not None:
                self.thread.request_cancel()
                self.status.setText(self._tr("取消中...", "Cancelling..."))
                self.btn_cancel.setEnabled(False)
        except Exception as e:
            QMessageBox.warning(self, self._tr("錯誤", "Error"), str(e))

    def _human_size(self, bytes_num):
        try:
            b = float(bytes_num)
        except Exception:
            return "-"
        if b <= 0:
            return "0 B"
        units = ['B','KB','MB','GB','TB']
        idx = 0
        while b >= 1024 and idx < len(units)-1:
            b /= 1024.0
            idx += 1
        return f"{b:.1f} {units[idx]}"

    def _format_eta(self, eta):
        if eta is None:
            return ""
        try:
            s = int(float(eta))
            if s <= 0:
                return ""
            h = s // 3600
            m = (s % 3600) // 60
            sec = s % 60
            if h > 0:
                return f"{h}h{m}m"
            if m > 0:
                return f"{m}m{sec}s"
            return f"{sec}s"
        except Exception:
            return ""

    def open_folder(self):
        try:
            folder = self.output_dir
            if folder:
                QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
        except Exception as e:
            QMessageBox.warning(self, self._tr("錯誤", "Error"), str(e))

    def open_file(self):
        try:
            fn = self.last_downloaded_file
            if fn and os.path.exists(fn):
                QDesktopServices.openUrl(QUrl.fromLocalFile(fn))
            else:
                QMessageBox.warning(self, self._tr("錯誤", "Error"), self._tr("檔案不存在", "File not found"))
        except Exception as e:
            QMessageBox.warning(self, self._tr("錯誤", "Error"), str(e))

    def change_language(self, idx):
        if idx == 0:
            self.label.setText("貼上 YouTube 連結:")
            self.filename_label.setText("檔案名稱:")
            self.filename_input.setPlaceholderText("留空時使用預設標題")
            self.mode_label.setText("選擇下載格式:")
            self.cookie_label.setText("瀏覽器 Cookie:")
            self.cookie_input_label.setText("手動 Cookie:")
            self.cookie_input.setPlaceholderText("貼上 Cookie（例如：SID=...; HSID=...）")
            self.preview_section_label.setText("影片資訊預覽:")
            self.btn_download.setText("下載")
            self.btn_open_folder.setText("開啟資料夾")
            self.btn_open_file.setText("開啟檔案")
            self.btn_website.setText("官方網站")
            self.btn_github.setText("GitHub")
            self.btn_settings.setText("設定")
            self.btn_add_queue.setText("加入佇列")
            self.btn_start_queue.setText("開始佇列")
            self.btn_clear_queue.setText("清空佇列")
            self.btn_cancel.setText("取消下載")
        else:
            self.label.setText("Paste YouTube URL:")
            self.filename_label.setText("File Name:")
            self.filename_input.setPlaceholderText("Leave blank to use the default title")
            self.mode_label.setText("Choose Download Format:")
            self.cookie_label.setText("Browser Cookies:")
            self.cookie_input_label.setText("Manual Cookie:")
            self.cookie_input.setPlaceholderText("Paste Cookie (example: SID=...; HSID=...)")
            self.preview_section_label.setText("Media Preview:")
            self.btn_download.setText("Download")
            self.btn_open_folder.setText("Open Folder")
            self.btn_open_file.setText("Open File")
            self.btn_website.setText("Official Website")
            self.btn_github.setText("GitHub")
            self.btn_settings.setText("Settings")
            self.btn_add_queue.setText("Add to Queue")
            self.btn_start_queue.setText("Start Queue")
            self.btn_clear_queue.setText("Clear Queue")
            self.btn_cancel.setText("Cancel")
        self._populate_format_combo()
        self._populate_cookie_combo()
        self._update_ffmpeg_status_label()
        self._update_cookie_input_visibility()
        self.btn_open_folder.setToolTip(self._tr("在檔案管理器中開啟輸出資料夾", "Open the output folder in File Explorer"))
        self.btn_open_file.setToolTip(self._tr("使用預設程式開啟下載的檔案", "Open the downloaded file with the default app"))
        self.btn_settings.setToolTip(self._tr("開啟設定面板", "Open settings panel"))
        self.startup_time_label.setToolTip(self._tr("程式啟動時間", "Application startup time"))
        if not self.startup_complete:
            self.status.setText(self._startup_status_text())
        elif not self.is_downloading and not self.preview_loading and not self.preview_data and not self.url_input.text().strip():
            self.status.setText(self._idle_status_text())
        self._refresh_preview_panel()
        self._update_button_states()
        self._persist_settings_if_ready()

    def on_cookie_source_changed(self):
        self._update_cookie_input_visibility()
        self._persist_settings_if_ready()
        current_url = self.url_input.text().strip()
        self.preview_data = None
        self.preview_url = None
        self._refresh_preview_panel()

        if not current_url or self.is_downloading or self.preview_loading:
            return

        self.auto_preview_timer.stop()
        self._start_preview_request(current_url)

    def on_manual_cookie_changed(self):
        if self.cookie_combo.currentData() != 'manual':
            return

        self._persist_settings_if_ready()
        current_url = self.url_input.text().strip()
        self.preview_data = None
        self.preview_url = None
        self._refresh_preview_panel()

        if not current_url or self.is_downloading or self.preview_loading:
            return

        self.auto_preview_timer.stop()
        self.auto_preview_timer.start(500)

    def on_error(self, msg):
        self.progress.setValue(0)
        self.progress_label.setText("0%")
        if isinstance(msg, str) and (msg.startswith('Cancelled') or 'cancel' in msg.lower()):
            self.status.setText(self._tr("已取消", "Cancelled"))
            QMessageBox.information(self, self._tr("已取消", "Cancelled"), msg)
        else:
            self.status.setText(self._tr("下載失敗", "Download Failed"))
            QMessageBox.critical(self, self._tr("錯誤", "Error"), msg)
        self.pending_queue_start = bool(self.queue)
    
    def open_website(self):
        url = QUrl("https://sites.google.com/view/yt-music-downloader/最新版本")
        QDesktopServices.openUrl(url)

    def open_github(self):
        url = QUrl("https://github.com/HubgaBro/YouTube-Video-Downloader")
        QDesktopServices.openUrl(url)

    def open_settings_dialog(self):
        dialog = SettingsDialog(
            self.COOKIE_BROWSER_OPTIONS,
            self.lang_combo.currentIndex(),
            self.output_dir,
            self.cookie_combo.currentData(),
            self.cookie_input.text(),
            self.app_version,
            self.ffmpeg_path,
            self,
        )
        dialog.website_button.clicked.connect(self.open_website)
        dialog.github_button.clicked.connect(self.open_github)

        if dialog.exec() != QDialog.Accepted:
            return

        settings = self._normalize_settings(dialog.selected_settings())
        self._apply_settings(settings)
        self._save_settings(settings, show_error=True)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    splash_pixmap = QPixmap(400, 150)
    splash_pixmap.fill(Qt.GlobalColor.white)
    
    icon_pixmap = QPixmap("YVD_ico.ico")
    if not icon_pixmap.isNull():
        icon_pixmap = icon_pixmap.scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        
        painter = QPainter(splash_pixmap)
        painter.drawPixmap(20, 20, icon_pixmap)
        painter.end()
    
    splash = QSplashScreen(splash_pixmap)
    splash.showMessage(
        "程式正在啟動...\nApplication is starting...",
        Qt.AlignCenter,
        Qt.GlobalColor.black,
    )
    splash.show()
    app.processEvents()

    window = MainWindow()
    window.show()
    app.processEvents()
    window.set_startup_time(time.perf_counter() - APP_START_TIME)
    splash.finish(window)
    sys.exit(app.exec())