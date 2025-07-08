import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QListWidget, QFileDialog, QStatusBar, QCheckBox, QGroupBox, QMenu, QTextEdit
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QEvent, QSettings
from PyQt5.QtGui import QFont, QCursor, QTextCharFormat, QTextCursor, QColor
from file_searcher import FileSearcher
import subprocess
import platform

class SearchThread(QThread):
    dosya_bulundu = pyqtSignal(str)
    arama_bitti = pyqtSignal(int)
    arama_durumu = pyqtSignal(str)

    def __init__(self, directory, keywords, extensions):
        super().__init__()
        self.directory = directory
        self.keywords = keywords
        self.extensions = extensions  # Liste: ['.txt', ...]
        self.file_searcher = FileSearcher()
        self._stop_requested = False

    def run(self):
        self.arama_durumu.emit("Arama yapılıyor...")
        keyword_list = [k.strip() for k in self.keywords.split(',') if k.strip()]
        if not keyword_list:
            self.arama_durumu.emit("Lütfen aranacak kelimeleri girin.")
            self.arama_bitti.emit(0)
            return
        if not self.extensions:
            self.arama_durumu.emit("Lütfen en az bir dosya türü seçin.")
            self.arama_bitti.emit(0)
            return
        bulunan_sayisi = 0
        try:
            for root, dirs, files in os.walk(self.directory):
                if self._stop_requested:
                    self.arama_durumu.emit("Arama iptal edildi.")
                    self.arama_bitti.emit(bulunan_sayisi)
                    return
                for file in files:
                    if self._stop_requested:
                        self.arama_durumu.emit("Arama iptal edildi.")
                        self.arama_bitti.emit(bulunan_sayisi)
                        return
                    file_path = os.path.join(root, file)
                    file_extension = os.path.splitext(file_path)[1].lower()
                    if file_extension not in self.extensions:
                        continue
                    try:
                        content = ""
                        if file_extension == '.txt':
                            with open(file_path, 'r', encoding='utf-8') as f:
                                content = f.read()
                        elif file_extension == '.docx':
                            from docx import Document
                            doc = Document(file_path)
                            content = '\n'.join([p.text for p in doc.paragraphs])
                        elif file_extension == '.pdf':
                            import PyPDF2
                            with open(file_path, 'rb') as f:
                                reader = PyPDF2.PdfReader(f)
                                content = '\n'.join([page.extract_text() or '' for page in reader.pages])
                        elif file_extension == '.xlsx':
                            from openpyxl import load_workbook
                            wb = load_workbook(file_path, data_only=True)
                            texts = []
                            for sheet in wb.worksheets:
                                for row in sheet.iter_rows(values_only=True):
                                    for cell in row:
                                        if cell is not None:
                                            texts.append(str(cell))
                            content = '\n'.join(texts)
                        # Anahtar kelime arama (büyük/küçük harf duyarsız)
                        content_lower = content.lower()
                        if any(kw.lower() in content_lower for kw in keyword_list):
                            self.dosya_bulundu.emit(file_path)
                            bulunan_sayisi += 1
                    except Exception as e:
                        continue
            self.arama_bitti.emit(bulunan_sayisi)
        except Exception as e:
            self.arama_durumu.emit(f"Arama sırasında hata oluştu: {str(e)}")
            self.arama_bitti.emit(bulunan_sayisi)

    def stop(self):
        self._stop_requested = True

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dosya İçeriği Arama Motoru")
        self.setGeometry(200, 200, 750, 700)
        self.search_thread = None
        self.settings = QSettings("Beyza", "DosyaAramaUygulamasi")
        self.init_ui()

    def init_ui(self):
        # Ana widget ve ana layout
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 30, 30, 20)
        main_layout.setSpacing(18)

        # Başlık
        title = QLabel("Dosya İçeriği Arama Motoru")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # Dizin seçimi
        dir_layout = QHBoxLayout()
        dir_label = QLabel("Aranacak Dizin:")
        dir_label.setMinimumWidth(120)
        self.dir_edit = QLineEdit()
        self.dir_edit.setReadOnly(True)
        self.dir_edit.setMinimumHeight(32)
        dir_btn = QPushButton("Dizin Seç")
        dir_btn.setMinimumHeight(32)
        dir_btn.clicked.connect(self.select_directory)
        dir_layout.addWidget(dir_label)
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(dir_btn)
        main_layout.addLayout(dir_layout)

        # Son kaydedilen dizini yükle
        last_dir = self.settings.value("last_directory", "")
        if last_dir:
            self.dir_edit.setText(last_dir)

        # Kelime girişi
        word_layout = QHBoxLayout()
        word_label = QLabel("Aranacak Kelimeler:")
        word_label.setMinimumWidth(120)
        self.word_edit = QLineEdit()
        self.word_edit.setMinimumHeight(32)
        self.word_edit.setPlaceholderText("Birden fazla kelime için virgül kullanın (örn: python, dosya)")
        word_layout.addWidget(word_label)
        word_layout.addWidget(self.word_edit)
        main_layout.addLayout(word_layout)

        # Dosya türü seçim kutuları grup kutusu içinde
        ext_group = QGroupBox("Dosya Türleri")
        ext_group.setStyleSheet("QGroupBox { font-weight: bold; border: 1px solid #bbb; border-radius: 6px; margin-top: 8px; } QGroupBox:title { subcontrol-origin: margin; left: 10px; padding:0 3px 0 3px; }")
        ext_layout = QHBoxLayout(ext_group)
        ext_layout.setContentsMargins(12, 8, 12, 8)
        self.cb_txt = QCheckBox(".txt")
        self.cb_docx = QCheckBox(".docx")
        self.cb_pdf = QCheckBox(".pdf")
        self.cb_xlsx = QCheckBox(".xlsx")
        for cb in [self.cb_txt, self.cb_docx, self.cb_pdf, self.cb_xlsx]:
            cb.setChecked(True)
            cb.setMinimumWidth(70)
            cb.setStyleSheet("font-size: 14px;")
        ext_layout.addWidget(self.cb_txt)
        ext_layout.addWidget(self.cb_docx)
        ext_layout.addWidget(self.cb_pdf)
        ext_layout.addWidget(self.cb_xlsx)
        main_layout.addWidget(ext_group)

        # Arama/Durdur butonu ve Sonuçları Kaydet butonu
        btn_layout = QHBoxLayout()
        self.search_btn = QPushButton("Aramayı Başlat")
        self.search_btn.setMinimumHeight(40)
        self.search_btn.setStyleSheet("background-color: #1976d2; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")
        self.search_btn.clicked.connect(self.toggle_search)
        btn_layout.addWidget(self.search_btn)
        self.save_btn = QPushButton("Sonuçları Kaydet")
        self.save_btn.setMinimumHeight(40)
        self.save_btn.setStyleSheet("background-color: #388e3c; color: white; font-weight: bold; font-size: 15px; border-radius: 8px;")
        self.save_btn.clicked.connect(self.save_results)
        btn_layout.addWidget(self.save_btn)
        main_layout.addLayout(btn_layout)

        # Sonuç listesi
        self.result_list = QListWidget()
        self.result_list.setStyleSheet("font-size: 14px; background: #f8f8f8; border: 1px solid #bbb; border-radius: 6px;")
        self.result_list.setMinimumHeight(180)
        main_layout.addWidget(self.result_list)
        self.result_list.itemDoubleClicked.connect(self.open_selected_file)
        self.result_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_list.customContextMenuRequested.connect(self.show_context_menu)
        self.result_list.currentItemChanged.connect(self.show_preview)

        # İçerik önizleme kutusu
        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setMinimumHeight(120)
        self.preview_box.setStyleSheet("font-size: 14px; background: #fff; border: 1px solid #bbb; border-radius: 6px; padding: 8px;")
        main_layout.addWidget(QLabel("İçerik Önizlemesi (aranan kelimeyi içeren satırlar):"))
        main_layout.addWidget(self.preview_box)

        # Durum çubuğu
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.setStyleSheet("font-size: 13px; color: #333;")
        self.status_bar.showMessage("Lütfen bir dizin seçin.")

        # Genel pencere arka planı
        self.setStyleSheet("QWidget { background: #f4f6fa; } QLabel { font-size: 15px; }")

        self._searching = False

    def select_directory(self):
        folder = QFileDialog.getExistingDirectory(self, "Dizin Seç")
        if folder:
            self.dir_edit.setText(folder)
            self.status_bar.showMessage("Dizin seçildi. Aranacak kelimeleri girin.")
            # Ayarlara kaydet
            self.settings.setValue("last_directory", folder)

    def toggle_search(self):
        if not self._searching:
            self.start_search()
        else:
            self.stop_search()

    def start_search(self):
        directory = self.dir_edit.text().strip()
        keywords = self.word_edit.text().strip()
        extensions = []
        if self.cb_txt.isChecked():
            extensions.append('.txt')
        if self.cb_docx.isChecked():
            extensions.append('.docx')
        if self.cb_pdf.isChecked():
            extensions.append('.pdf')
        if self.cb_xlsx.isChecked():
            extensions.append('.xlsx')
        if not directory:
            self.status_bar.showMessage("Lütfen bir dizin seçin.")
            return
        if not keywords:
            self.status_bar.showMessage("Lütfen aranacak kelimeleri girin.")
            return
        if not extensions:
            self.status_bar.showMessage("Lütfen en az bir dosya türü seçin.")
            return
        self.result_list.clear()
        self.preview_box.clear()
        self.status_bar.showMessage("Arama yapılıyor...")
        self.search_btn.setText("Aramayı Durdur")
        self.search_btn.setStyleSheet("background-color: #d32f2f; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")
        self.search_btn.setEnabled(True)
        self._searching = True
        self.search_thread = SearchThread(directory, keywords, extensions)
        self.search_thread.dosya_bulundu.connect(self.add_result)
        self.search_thread.arama_bitti.connect(self.search_finished)
        self.search_thread.arama_durumu.connect(self.status_bar.showMessage)
        self.search_thread.start()

    def stop_search(self):
        if self.search_thread and self._searching:
            self.search_thread.stop()
            self.status_bar.showMessage("Arama iptal ediliyor...")
            self.search_btn.setEnabled(False)

    def add_result(self, file_path):
        self.result_list.addItem(file_path)

    def search_finished(self, count):
        if self._searching:
            self.status_bar.showMessage(f"Arama tamamlandı. {count} dosya bulundu.")
        else:
            self.status_bar.showMessage("Arama iptal edildi.")
        self.search_btn.setText("Aramayı Başlat")
        self.search_btn.setStyleSheet("background-color: #1976d2; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")
        self.search_btn.setEnabled(True)
        self._searching = False

    def open_selected_file(self, item):
        file_path = item.text()
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", file_path])
            else:
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            self.status_bar.showMessage(f"Dosya açılamadı: {e}")

    def show_context_menu(self, pos):
        item = self.result_list.itemAt(pos)
        if not item:
            return
        file_path = item.text()
        menu = QMenu()
        ac_action = menu.addAction("Dosyayı Aç")
        konum_action = menu.addAction("Dosyanın Konumunu Aç")
        kopyala_action = menu.addAction("Yolu Kopyala")
        action = menu.exec_(QCursor.pos())
        if action == ac_action:
            self.open_selected_file(item)
        elif action == konum_action:
            self.open_file_location(file_path)
        elif action == kopyala_action:
            self.copy_file_path(file_path)

    def open_file_location(self, file_path):
        try:
            folder = os.path.dirname(file_path)
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer /select,"{file_path}"')
            elif platform.system() == "Darwin":
                subprocess.call(["open", folder])
            else:
                subprocess.call(["xdg-open", folder])
        except Exception as e:
            self.status_bar.showMessage(f"Konum açılamadı: {e}")

    def copy_file_path(self, file_path):
        try:
            cb = QApplication.clipboard()
            cb.setText(file_path)
            self.status_bar.showMessage("Dosya yolu panoya kopyalandı.")
        except Exception as e:
            self.status_bar.showMessage(f"Kopyalama hatası: {e}")

    def save_results(self):
        if self.result_list.count() == 0:
            self.status_bar.showMessage("Kaydedilecek sonuç yok.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Sonuçları Kaydet", "", "Metin Dosyası (*.txt);;CSV Dosyası (*.csv)")
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                for i in range(self.result_list.count()):
                    line = self.result_list.item(i).text()
                    if path.endswith('.csv'):
                        f.write(f'"{line}"\n')
                    else:
                        f.write(line + '\n')
            self.status_bar.showMessage(f"Sonuçlar kaydedildi: {path}")
        except Exception as e:
            self.status_bar.showMessage(f"Kayıt hatası: {e}")

    def show_preview(self, current, previous):
        self.preview_box.clear()
        if not current:
            return
        file_path = current.text()
        keywords = [k.strip() for k in self.word_edit.text().split(',') if k.strip()]
        if not keywords:
            return
        ext = os.path.splitext(file_path)[1].lower()
        try:
            content = ""
            if ext == '.txt':
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                matches = [line.strip() for line in lines if any(kw.lower() in line.lower() for kw in keywords)]
                content = '\n'.join(matches)
            elif ext == '.docx':
                from docx import Document
                doc = Document(file_path)
                matches = []
                for p in doc.paragraphs:
                    if any(kw.lower() in p.text.lower() for kw in keywords):
                        matches.append(p.text.strip())
                content = '\n'.join(matches)
            elif ext == '.pdf':
                import PyPDF2
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    matches = []
                    for page in reader.pages:
                        text = page.extract_text() or ''
                        for line in text.splitlines():
                            if any(kw.lower() in line.lower() for kw in keywords):
                                matches.append(line.strip())
                content = '\n'.join(matches)
            elif ext == '.xlsx':
                from openpyxl import load_workbook
                wb = load_workbook(file_path, data_only=True)
                matches = []
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        for cell in row:
                            if cell is not None and any(kw.lower() in str(cell).lower() for kw in keywords):
                                matches.append(str(cell))
                content = '\n'.join(matches)
            if not content:
                self.preview_box.setPlainText("Aranan kelimeyi içeren satır veya paragraf bulunamadı.")
            else:
                self.preview_box.setPlainText(content)
                # Vurgulama
                self.highlight_keywords(keywords)
        except Exception as e:
            self.preview_box.setPlainText(f"Önizleme hatası: {e}")

    def highlight_keywords(self, keywords):
        cursor = self.preview_box.textCursor()
        fmt = QTextCharFormat()
        fmt.setBackground(QColor("yellow"))
        text = self.preview_box.toPlainText().lower()
        for kw in keywords:
            if not kw:
                continue
            cursor.setPosition(0)
            while True:
                idx = text.find(kw.lower(), cursor.position())
                if idx == -1:
                    break
                cursor.setPosition(idx)
                cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, len(kw))
                cursor.mergeCharFormat(fmt)
                cursor.setPosition(idx + len(kw))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 