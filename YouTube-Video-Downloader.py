import sys
import yt_dlp
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, 
    QProgressBar, QMessageBox, QFileDialog
)
from PySide6.QtCore import QThread, Signal, QUrl, Qt
from PySide6.QtGui import QDesktopServices

class DownloadThread(QThread):
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, url, format_selector, output_path, ffmpeg_location=None):
        super().__init__()
        self.url = url
        self.format_selector = format_selector
        self.output_path = output_path
        self.ffmpeg_location = ffmpeg_location

    def run(self):
        def progress_hook(d):
            if d.get('status') == 'downloading':
                try:
                        percent_str = d.get('_percent_str', '')
                        if percent_str and '%' in percent_str:
                            p = float(percent_str.strip('%'))
                        else:
                            downloaded = d.get('downloaded_bytes', 0)
                            total = d.get('total_bytes', 0) or d.get('total_bytes_estimate', 0)
                            if total > 0:
                                p = (downloaded / total) * 100
                            else:
                                p = 0
                        p = min(p, 100)
                        self.progress.emit(int(p))
                except:
                        pass

        common_opts = {
            'outtmpl': f'{self.output_path}/%(title)s.%(ext)s',
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

        if "MP4 | High Resolution" in self.format_selector:
            ydl_opts = {
                **common_opts,
                'format': 'bestvideo+bestaudio/best',
                'merge_output_format': 'mp4',
                'postprocessor_args': {'FFmpegMerger': ['-movflags', '+faststart']},
            }
        elif "MP3 | high" in self.format_selector:
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '0',
                }],
            }
        elif "MP3 | low" in self.format_selector:
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '5',
                }],
            }
        elif "WAV" in self.format_selector:
            ydl_opts = {
                **common_opts,
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'wav',
                }],
            }
        elif "OGG" in self.format_selector:
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
                ydl.download([self.url])
            self.finished.emit("Download and Conversion Complete!")
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YouTube Video Downloader     1.0")
        self.setFixedSize(650, 420)

        self.status = QLabel("Download Ready. . .")
        self.ffmpeg_path = self.find_ffmpeg()

        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        layout.setSpacing(8)
        layout.setContentsMargins(25, 25, 25, 25)

        self.label = QLabel("Paste YouTube URL:")
        layout.addWidget(self.label)

        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://www.youtube.com/watch?v=...")
        layout.addWidget(self.url_input)

        self.mode_label = QLabel("Choose Download Format:")
        layout.addWidget(self.mode_label)

        self.format_combo = QComboBox()
        self.format_combo.addItems([
            "MP4 | High Resolution",
            "MP3 | high kbit/s",
            "MP3 | low kbit/s",
            "WAV | non-Lossy",
            "OGG | Lossy"
        ])
        layout.addWidget(self.format_combo)

        self.btn_download = QPushButton("Download")
        self.btn_download.clicked.connect(self.start_download)
        layout.addWidget(self.btn_download)

        progress_layout = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setMaximum(100)
        progress_layout.addWidget(self.progress)
        
        self.progress_label = QLabel("0%")
        self.progress_label.setMinimumWidth(40)
        progress_layout.addWidget(self.progress_label)
        layout.addLayout(progress_layout)

        layout.addWidget(self.status)
        
        bottom_layout = QHBoxLayout()
        self.btn_website = QPushButton("Official Website")
        self.btn_website.clicked.connect(self.open_website)
        bottom_layout.addWidget(self.btn_website)
        self.btn_github = QPushButton("GitHub")
        self.btn_github.clicked.connect(self.open_github)
        bottom_layout.addWidget(self.btn_github)
        bottom_layout.addStretch()
        layout.addLayout(bottom_layout)

        self.output_dir = "下載"

    def find_ffmpeg(self):
        current_dir = Path(__file__).parent.absolute()
        
        ffmpeg_exe = current_dir / "ffmpeg.exe"
        if ffmpeg_exe.exists():
            self.status.setText("Download Ready. . . (ffmpeg 已自動載入)")
            return str(ffmpeg_exe)
        
        ffmpeg_bin = current_dir / "ffmpeg"
        if ffmpeg_bin.exists():
            self.status.setText("Download Ready. . . (ffmpeg 已自動載入)")
            return str(ffmpeg_bin)
        
        self.status.setText("Download Ready. . . (未找到 ffmpeg，轉檔功能可能失效)")
        return None

    def start_download(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "Error", "Please enter a YouTube URL")
            return

        selected = self.format_combo.currentText()
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory", self.output_dir)
        if not folder:
            return
        self.output_dir = folder

        self.progress.setValue(0)
        self.btn_download.setEnabled(False)
        self.status.setText(f"Downloading. . . ({selected})")

        self.thread = DownloadThread(url, selected, folder, self.ffmpeg_path)
        self.thread.progress.connect(self.update_progress, Qt.QueuedConnection)
        self.thread.finished.connect(self.on_finished)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_finished(self, msg):
        self.progress.setValue(100)
        self.progress_label.setText("100%")
        self.status.setText(msg)
        QMessageBox.information(self, "Done", msg)
        self.btn_download.setEnabled(True)

    def update_progress(self, value):
        """更新進度條和百分比標籤"""
        self.progress.setValue(value)
        self.progress_label.setText(f"{value}%")

    def on_error(self, msg):
        self.progress.setValue(0)
        self.progress_label.setText("0%")
        self.status.setText("Download Failed")
        QMessageBox.critical(self, "錯誤", msg)
        self.btn_download.setEnabled(True)
    
    def open_website(self):
        url = QUrl("https://sites.google.com/view/yt-music-downloader/最新版本")
        QDesktopServices.openUrl(url)

    def open_github(self):
        url = QUrl("https://github.com/HubgaBro/YouTube-Video-Downloader")
        QDesktopServices.openUrl(url)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())