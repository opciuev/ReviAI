"""
ReviAI - AI評審作業システム
PySide6 GUI Main Application
"""
import warnings
# Suppress warnings from dependencies
warnings.filterwarnings("ignore", message="Unsupported Windows version")
warnings.filterwarnings("ignore", message="Couldn't find ffmpeg or avconv")
warnings.filterwarnings("ignore", category=UserWarning, module="onnxruntime")
warnings.filterwarnings("ignore", category=RuntimeWarning, module="pydub")

import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget,
    QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
    QLineEdit, QPushButton, QTextEdit, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox, QScrollArea,
    QDialog, QDialogButtonBox, QComboBox, QListWidget, QListWidgetItem,
    QInputDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QUrl
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent

# Import modules
from config_manager import ConfigManager
from models import ReviewTable
from logger import logger
import step1_excel_to_pdf as step1
import step2_ai_review as step2
import step3_save_results as step3


# Custom drag-drop widgets
class DragDropLineEdit(QLineEdit):
    """LineEdit with drag and drop support for files"""
    def __init__(self, parent=None, file_filter="*"):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_filter = file_filter  # e.g., ".xlsx", ".pdf"
        self.setPlaceholderText("ファイルをドラッグ＆ドロップまたはクリックして選択")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            # Filter by extension if specified
            if self.file_filter != "*":
                valid_files = [f for f in files if f.lower().endswith(self.file_filter.lower())]
                if valid_files:
                    self.setText(valid_files[0])  # Use first valid file
                else:
                    QMessageBox.warning(
                        self,
                        "エラー",
                        f"対応していないファイル形式です\n\n対応形式: {self.file_filter}"
                    )
            else:
                self.setText(files[0])


class DragDropListWidget(QListWidget):
    """ListWidget with drag and drop support for files"""
    def __init__(self, parent=None, file_filter="*"):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_filter = file_filter  # e.g., ".pdf", or list like [".pdf", ".md"]
        self.parent_widget = parent

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            # Filter by extension
            if self.file_filter != "*":
                # Support both single string and list of extensions
                if isinstance(self.file_filter, str):
                    valid_files = [f for f in files if f.lower().endswith(self.file_filter.lower())]
                    supported_formats = self.file_filter
                else:  # List of extensions
                    valid_files = [f for f in files
                                   if any(f.lower().endswith(ext.lower()) for ext in self.file_filter)]
                    supported_formats = ", ".join(self.file_filter)
            else:
                valid_files = files
                supported_formats = "*"

            if valid_files:
                # Add valid files to parent's file list
                if hasattr(self.parent_widget, 'add_dropped_files'):
                    self.parent_widget.add_dropped_files(valid_files)
            else:
                QMessageBox.warning(
                    self,
                    "エラー",
                    f"対応していないファイル形式です\n\n対応形式: {supported_formats}"
                )


class SheetLoaderWorker(QThread):
    """Worker thread for loading Excel sheets to prevent UI freezing"""
    finished = Signal(list)
    error = Signal(str)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        try:
            sheets = step1.list_all_sheets(self.excel_path)
            self.finished.emit(sheets)
        except Exception as e:
            self.error.emit(str(e))


class PDFGeneratorWorker(QThread):
    """Worker thread for PDF generation to prevent UI freezing"""
    finished = Signal(list)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, excel_path, sheet_names, version, output_dir):
        super().__init__()
        self.excel_path = excel_path
        self.sheet_names = sheet_names
        self.version = version
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("PDF生成中...")
            pdf_files = step1.generate_pdfs(
                self.excel_path,
                self.sheet_names,
                self.version,
                self.output_dir
            )
            self.finished.emit(pdf_files)
        except Exception as e:
            self.error.emit(str(e))


class AIReviewWorker(QThread):
    """Worker thread for AI review to prevent UI freezing"""
    finished = Signal(ReviewTable)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, pdf_paths, prompt, api_key, model):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.prompt = prompt
        self.api_key = api_key
        self.model = model

    def run(self):
        try:
            self.progress.emit("AI評審を開始しています...")
            result = step2.review_with_retry(
                self.pdf_paths,
                self.prompt,
                self.api_key,
                self.model,
                max_retries=3
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class APIKeyDialog(QDialog):
    """Dialog for managing API Key"""
    def __init__(self, parent, current_key=""):
        super().__init__(parent)
        self.setWindowTitle("API Key 設定")
        self.setMinimumSize(500, 200)

        layout = QVBoxLayout()

        # Info label
        info_label = QLabel(
            "Gemini API Keyを入力してください\n"
            "API Keyをお持ちでない場合は、下のボタンから取得できます"
        )
        info_label.setStyleSheet("color: #666; padding: 10px;")
        layout.addWidget(info_label)

        # API Key input
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("API Key:"))

        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("API Keyを入力してください")
        self.key_input.setText(current_key)
        self.key_input.setEchoMode(QLineEdit.Password)
        key_layout.addWidget(self.key_input, stretch=1)

        # Show/Hide button
        self.show_btn = QPushButton("表示")
        self.show_btn.setMaximumWidth(60)
        self.show_btn.setCheckable(True)
        self.show_btn.toggled.connect(self.toggle_visibility)
        key_layout.addWidget(self.show_btn)

        layout.addLayout(key_layout)

        # Help button
        help_btn = QPushButton("API Keyを取得する (ブラウザで開く)")
        help_btn.clicked.connect(self.open_help)
        layout.addWidget(help_btn)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def toggle_visibility(self, checked):
        """Toggle API key visibility"""
        if checked:
            self.key_input.setEchoMode(QLineEdit.Normal)
        else:
            self.key_input.setEchoMode(QLineEdit.Password)

    def open_help(self):
        """Open API key help URL"""
        import webbrowser
        webbrowser.open("https://aistudio.google.com/apikey")

    def get_api_key(self):
        """Get entered API key"""
        return self.key_input.text().strip()


class PromptManagerDialog(QDialog):
    """Dialog for managing multiple prompt templates"""
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("プロンプトテンプレート管理")
        self.setMinimumSize(900, 600)

        self.prompts_dir = Path("./prompts")
        self.prompts_dir.mkdir(exist_ok=True)
        self.selected_prompt = None

        self.init_ui()
        self.load_prompt_list()

    def init_ui(self):
        layout = QHBoxLayout()

        # Left panel: List of prompts
        left_panel = QVBoxLayout()

        left_panel.addWidget(QLabel("保存済みプロンプトテンプレート:"))

        self.prompt_list = QListWidget()
        self.prompt_list.setMaximumWidth(300)
        self.prompt_list.itemClicked.connect(self.on_prompt_selected)
        self.prompt_list.itemDoubleClicked.connect(self.edit_prompt)
        left_panel.addWidget(self.prompt_list)

        # Buttons for list management
        btn_layout = QVBoxLayout()

        new_btn = QPushButton("新規作成")
        new_btn.clicked.connect(self.create_new_prompt)
        btn_layout.addWidget(new_btn)

        edit_btn = QPushButton("編集")
        edit_btn.clicked.connect(self.edit_prompt)
        btn_layout.addWidget(edit_btn)

        rename_btn = QPushButton("名前変更")
        rename_btn.clicked.connect(self.rename_prompt)
        btn_layout.addWidget(rename_btn)

        copy_btn = QPushButton("複製")
        copy_btn.clicked.connect(self.copy_prompt)
        btn_layout.addWidget(copy_btn)

        delete_btn = QPushButton("削除")
        delete_btn.clicked.connect(self.delete_prompt)
        btn_layout.addWidget(delete_btn)

        btn_layout.addStretch()
        left_panel.addLayout(btn_layout)

        layout.addLayout(left_panel)

        # Right panel: Preview
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("プレビュー:"))

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Consolas", 10))
        right_panel.addWidget(self.preview_text)

        layout.addLayout(right_panel)

        # Bottom buttons
        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        self.setLayout(main_layout)

    def load_prompt_list(self):
        """Load all saved prompts"""
        self.prompt_list.clear()

        # Load all prompts from prompts directory
        prompt_files = sorted(self.prompts_dir.glob("*.txt"))

        if not prompt_files:
            # Create a default prompt if none exists
            default_prompt = """# AI評審プロンプト

このファイルにプロンプトを記入してください。"""
            default_file = self.prompts_dir / "標準テンプレート.txt"
            with open(default_file, 'w', encoding='utf-8') as f:
                f.write(default_prompt)
            prompt_files = [default_file]
            logger.info(f"Created default prompt: {default_file}")

        for prompt_file in prompt_files:
            item = QListWidgetItem(prompt_file.stem)
            item.setData(Qt.UserRole, str(prompt_file))
            self.prompt_list.addItem(item)

    def on_prompt_selected(self, item):
        """Preview selected prompt"""
        prompt_path = item.data(Qt.UserRole)

        try:
            with open(prompt_path, 'r', encoding='utf-8') as f:
                content = f.read()

            self.preview_text.setText(content)
            self.selected_prompt = prompt_path
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"プロンプト読込失敗:\n{str(e)}")

    def create_new_prompt(self):
        """Create new prompt template"""
        name, ok = QInputDialog.getText(
            self,
            "新規プロンプト作成",
            "プロンプト名を入力してください:"
        )

        if ok and name:
            # Open editor with empty content
            dialog = PromptEditorDialog(self, "")
            if dialog.exec() == QDialog.Accepted:
                content = dialog.get_prompt()
                if content.strip():
                    # Save to prompts directory
                    file_path = self.prompts_dir / f"{name}.txt"
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)

                    self.load_prompt_list()
                    QMessageBox.information(self, "成功", f"プロンプト '{name}' を作成しました")

    def edit_prompt(self):
        """Edit selected prompt"""
        current_item = self.prompt_list.currentItem()
        if not current_item:
            QMessageBox.information(self, "情報", "編集するプロンプトを選択してください")
            return

        prompt_path = current_item.data(Qt.UserRole)

        try:
            # Load current content
            with open(prompt_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Open editor
            dialog = PromptEditorDialog(self, content)
            if dialog.exec() == QDialog.Accepted:
                new_content = dialog.get_prompt()
                if new_content.strip():
                    with open(prompt_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)

                    self.preview_text.setText(new_content)
                    QMessageBox.information(self, "成功", "プロンプトを更新しました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"プロンプト編集失敗:\n{str(e)}")

    def rename_prompt(self):
        """Rename selected prompt"""
        current_item = self.prompt_list.currentItem()
        if not current_item:
            QMessageBox.information(self, "情報", "名前変更するプロンプトを選択してください")
            return

        old_path = Path(current_item.data(Qt.UserRole))
        old_name = old_path.stem

        name, ok = QInputDialog.getText(
            self,
            "プロンプト名前変更",
            "新しいプロンプト名を入力してください:",
            text=old_name
        )

        if ok and name and name != old_name:
            try:
                new_file = self.prompts_dir / f"{name}.txt"
                if new_file.exists():
                    QMessageBox.warning(self, "エラー", f"プロンプト '{name}' は既に存在します")
                    return

                old_path.rename(new_file)

                # Update current selection
                if self.selected_prompt == str(old_path):
                    self.selected_prompt = str(new_file)

                # Reload list and clear preview
                self.load_prompt_list()
                self.preview_text.clear()

                # Select the renamed item
                for i in range(self.prompt_list.count()):
                    item = self.prompt_list.item(i)
                    if item.text() == name:
                        self.prompt_list.setCurrentItem(item)
                        self.on_prompt_selected(item)
                        break

                QMessageBox.information(self, "成功", f"プロンプト名を '{name}' に変更しました")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"名前変更失敗:\n{str(e)}")

    def copy_prompt(self):
        """Copy selected prompt"""
        current_item = self.prompt_list.currentItem()
        if not current_item:
            QMessageBox.information(self, "情報", "複製するプロンプトを選択してください")
            return

        name, ok = QInputDialog.getText(
            self,
            "プロンプト複製",
            "新しいプロンプト名を入力してください:"
        )

        if ok and name:
            prompt_path = current_item.data(Qt.UserRole)

            try:
                # Load source content
                with open(prompt_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                # Save with new name
                new_file = self.prompts_dir / f"{name}.txt"
                with open(new_file, 'w', encoding='utf-8') as f:
                    f.write(content)

                self.load_prompt_list()
                QMessageBox.information(self, "成功", f"プロンプト '{name}' を作成しました")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"複製失敗:\n{str(e)}")

    def delete_prompt(self):
        """Delete selected prompt"""
        current_item = self.prompt_list.currentItem()
        if not current_item:
            QMessageBox.information(self, "情報", "削除するプロンプトを選択してください")
            return

        prompt_path = current_item.data(Qt.UserRole)

        # Check if this is the last prompt
        prompt_count = len(list(self.prompts_dir.glob("*.txt")))
        if prompt_count <= 1:
            QMessageBox.warning(
                self,
                "警告",
                f"最後のプロンプトは削除できません\n\n現在 {prompt_count} 個のプロンプトがあります。\n少なくとも1つのプロンプトが必要です。\n\n他のプロンプトを作成してから削除してください。"
            )
            return

        reply = QMessageBox.question(
            self,
            "確認",
            f"プロンプト '{current_item.text()}' を削除しますか？\n\n(残り {prompt_count - 1} 個のプロンプト)",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                Path(prompt_path).unlink()

                # Clear selection
                self.selected_prompt = None

                # Reload list
                self.load_prompt_list()
                self.preview_text.clear()

                # Auto-select first item
                if self.prompt_list.count() > 0:
                    first_item = self.prompt_list.item(0)
                    self.prompt_list.setCurrentItem(first_item)
                    self.on_prompt_selected(first_item)

                QMessageBox.information(self, "成功", "プロンプトを削除しました")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"削除失敗:\n{str(e)}")

    def get_selected_prompt_path(self):
        """Get selected prompt file path"""
        return self.selected_prompt


class PromptEditorDialog(QDialog):
    """Dialog for editing prompt template"""
    def __init__(self, parent, current_prompt):
        super().__init__(parent)
        self.setWindowTitle("プロンプト編集")
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout()

        # Info label
        info_label = QLabel(
            "[ヒント] Ctrl+A (全選択) → 貼り付けで既存内容を置換できます\n"
            "保存時に余分な空行は自動的に削除されます"
        )
        info_label.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        layout.addWidget(info_label)

        # Text editor with better settings
        layout.addWidget(QLabel("評審プロンプト:"))
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_prompt)
        self.text_edit.setLineWrapMode(QTextEdit.NoWrap)  # Disable auto wrap
        self.text_edit.setTabStopDistance(40)  # Set tab width

        # Set monospace font for better formatting visibility
        from PySide6.QtGui import QFont
        font = QFont("Consolas", 10)
        if not font.exactMatch():
            font = QFont("Courier New", 10)
        self.text_edit.setFont(font)

        layout.addWidget(self.text_edit)

        # Character count
        self.char_count_label = QLabel("")
        self.update_char_count()
        self.text_edit.textChanged.connect(self.update_char_count)
        layout.addWidget(self.char_count_label)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def update_char_count(self):
        """Update character count"""
        text = self.text_edit.toPlainText()
        lines = len(text.split('\n'))
        chars = len(text)
        self.char_count_label.setText(f"行数: {lines} | 文字数: {chars}")

    def get_prompt(self):
        """Get prompt text with normalized line breaks"""
        text = self.text_edit.toPlainText()
        # Normalize line breaks: remove excessive empty lines (more than 2 consecutive)
        import re
        # Replace 3+ consecutive newlines with just 2
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text


class Step1Page(QWidget):
    """Step 1: Excel to PDF generation"""
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.excel_path = None
        self.sheet_checkboxes = []
        self.sheet_loader_worker = None
        self.pdf_generator_worker = None
        self.generated_pdf_files = []

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Title
        title = QLabel("第1段階: Excel→PDF生成")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Form
        form = QFormLayout()

        # File selection with drag-drop support
        file_layout = QHBoxLayout()
        self.file_input = DragDropLineEdit(file_filter=".xlsx")
        self.file_input.setReadOnly(True)
        self.file_input.mousePressEvent = lambda e: self.browse_excel()
        self.file_input.setStyleSheet("QLineEdit { background-color: white; }")
        file_layout.addWidget(self.file_input, stretch=1)

        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.browse_excel)
        file_layout.addWidget(browse_btn)

        clear_btn = QPushButton("クリア")
        clear_btn.clicked.connect(self.clear_all)
        file_layout.addWidget(clear_btn)

        form.addRow("Excelファイル:", file_layout)

        # Connect textChanged signal to handle dropped files
        self.file_input.textChanged.connect(self.on_excel_file_changed)

        # Sheet selection area (scrollable)
        self.sheet_scroll = QScrollArea()
        self.sheet_scroll.setMaximumHeight(200)
        self.sheet_widget = QWidget()
        self.sheet_layout = QVBoxLayout(self.sheet_widget)
        self.sheet_scroll.setWidget(self.sheet_widget)
        self.sheet_scroll.setWidgetResizable(True)

        # Loading indicator for sheets
        self.sheet_loading_label = QLabel("")
        self.sheet_layout.addWidget(self.sheet_loading_label)

        form.addRow("対象シート:", self.sheet_scroll)

        # Version input
        self.version_input = QLineEdit()
        self.version_input.setPlaceholderText("例: 6")
        form.addRow("バージョン番号:", self.version_input)

        layout.addLayout(form)

        # Generate button
        self.generate_btn = QPushButton("PDF生成 ▶")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.clicked.connect(self.generate_pdfs)
        layout.addWidget(self.generate_btn)

        # Progress bar for PDF generation
        self.pdf_progress = QProgressBar()
        self.pdf_progress.setVisible(False)
        layout.addWidget(self.pdf_progress)

        # Status label for PDF generation (simple text)
        self.pdf_status_label = QLabel("")
        self.pdf_status_label.setWordWrap(True)
        layout.addWidget(self.pdf_status_label)

        # Next button
        next_btn = QPushButton("次へ →")
        next_btn.clicked.connect(self.controller.next_step)
        layout.addWidget(next_btn, alignment=Qt.AlignRight)

        self.setLayout(layout)

    def on_excel_file_changed(self):
        """Handle Excel file selection (from drag-drop or browse)"""
        file_path = self.file_input.text()
        if file_path and Path(file_path).exists():
            self.excel_path = file_path
            self.load_sheets()
            logger.info(f"Excel file selected: {Path(file_path).name}")

    def browse_excel(self):
        """Open file dialog to select Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Excelファイルを選択",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_input.setText(file_path)

    def load_sheets(self):
        """Load and display all sheets from Excel file using background thread"""
        # Clear old checkboxes
        for cb in self.sheet_checkboxes:
            cb.deleteLater()
        self.sheet_checkboxes.clear()

        # Show loading indicator
        self.sheet_loading_label.setText("[読込中] シートを読み込み中...")
        self.sheet_loading_label.setStyleSheet("color: blue;")

        # Create worker thread
        self.sheet_loader_worker = SheetLoaderWorker(self.excel_path)
        self.sheet_loader_worker.finished.connect(self.on_sheets_loaded)
        self.sheet_loader_worker.error.connect(self.on_sheets_load_error)
        self.sheet_loader_worker.start()

    def on_sheets_loaded(self, sheets):
        """Handle successful sheet loading"""
        # Hide loading indicator
        self.sheet_loading_label.setText("")

        # Create checkboxes
        for sheet_name in sheets:
            cb = QCheckBox(sheet_name)
            self.sheet_layout.addWidget(cb)
            self.sheet_checkboxes.append(cb)

        logger.info(f"Loaded {len(sheets)} sheets")

    def on_sheets_load_error(self, error_msg):
        """Handle sheet loading error"""
        self.sheet_loading_label.setText("")
        QMessageBox.critical(self, "エラー", f"シート読み込み失敗:\n{error_msg}")

    def clear_all(self):
        """Clear all inputs and reset to initial state"""
        reply = QMessageBox.question(
            self,
            "確認",
            "すべての入力をクリアしますか？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Clear file selection
            self.excel_path = None
            self.file_input.clear()

            # Clear sheet checkboxes
            for cb in self.sheet_checkboxes:
                cb.deleteLater()
            self.sheet_checkboxes.clear()
            self.sheet_loading_label.setText("")

            # Clear version input
            self.version_input.clear()

            # Clear status label
            self.pdf_status_label.clear()

            # Clear generated PDF files
            self.generated_pdf_files.clear()

            logger.info("Step 1 cleared")

    def generate_pdfs(self):
        """Generate PDF files from selected sheets using background thread"""
        # Validate inputs
        if not self.excel_path:
            QMessageBox.warning(self, "エラー", "Excelファイルを選択してください")
            return

        selected_sheets = [
            cb.text() for cb in self.sheet_checkboxes if cb.isChecked()
        ]
        if not selected_sheets:
            QMessageBox.warning(self, "エラー", "シートを選択してください")
            return

        version = self.version_input.text()
        if not version.isdigit():
            QMessageBox.warning(self, "エラー", "バージョン番号は数字で入力してください")
            return

        # Disable button and show progress
        self.generate_btn.setEnabled(False)
        self.pdf_progress.setVisible(True)
        self.pdf_progress.setRange(0, 0)  # Indeterminate mode
        self.pdf_status_label.setText("PDF生成中...")
        self.pdf_status_label.setStyleSheet("color: blue;")

        # Generate PDFs in background thread
        output_dir = "./output/pdfs"
        self.pdf_generator_worker = PDFGeneratorWorker(
            self.excel_path,
            selected_sheets,
            int(version),
            output_dir
        )
        self.pdf_generator_worker.finished.connect(self.on_pdfs_generated)
        self.pdf_generator_worker.error.connect(self.on_pdf_generation_error)
        self.pdf_generator_worker.progress.connect(self.on_pdf_progress)
        self.pdf_generator_worker.start()

    def on_pdf_progress(self, message):
        """Update PDF generation progress"""
        self.pdf_status_label.setText(message)

    def on_pdfs_generated(self, pdf_files):
        """Handle successful PDF generation"""
        # Reset UI
        self.generate_btn.setEnabled(True)
        self.pdf_progress.setVisible(False)

        # Save generated files
        self.generated_pdf_files = pdf_files

        # Save version number to controller for Step3
        version = self.version_input.text()
        if version.isdigit():
            self.controller.version_number = int(version)
            logger.info(f"Saved version number for Step3: {version}")

        # Display simple success message
        self.pdf_status_label.setText(f"[完了] PDF生成完了 ({len(pdf_files)}個)")
        self.pdf_status_label.setStyleSheet("color: green; font-weight: bold;")

        # Save paths for next step
        self.controller.step2_page.set_pdf_files(pdf_files)

        # Ask user if they want to open the PDFs
        reply = QMessageBox.question(
            self,
            "完了",
            f"{len(pdf_files)}個のPDFファイルを生成しました\n\n生成されたPDFファイルを開きますか？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.open_generated_pdfs()

    def on_pdf_generation_error(self, error_msg):
        """Handle PDF generation error"""
        self.generate_btn.setEnabled(True)
        self.pdf_progress.setVisible(False)
        self.pdf_status_label.setText("[失敗] PDF生成失敗")
        self.pdf_status_label.setStyleSheet("color: red; font-weight: bold;")
        QMessageBox.critical(self, "エラー", f"PDF生成失敗:\n{error_msg}")
        logger.error(f"PDF generation error: {error_msg}")

    def open_generated_pdfs(self):
        """Open generated PDF files in default viewer"""
        import os
        import subprocess

        for pdf_path in self.generated_pdf_files:
            try:
                # Use os.startfile on Windows to open with default application
                if sys.platform == "win32":
                    os.startfile(pdf_path)
                elif sys.platform == "darwin":  # macOS
                    subprocess.run(["open", pdf_path])
                else:  # Linux
                    subprocess.run(["xdg-open", pdf_path])

                logger.info(f"Opened PDF: {pdf_path}")
            except Exception as e:
                logger.error(f"Failed to open PDF {pdf_path}: {str(e)}")


class Step2Page(QWidget):
    """Step 2: AI Review"""
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.pdf_files = []
        self.review_result = None
        self.worker = None
        self.api_key = ""
        self.current_prompt_path = None  # Will auto-select first available prompt

        self.init_ui()
        self.load_api_key()
        self.load_prompts()  # Load prompt list after UI init

    def init_ui(self):
        layout = QVBoxLayout()

        # Title
        title = QLabel("第2段階: AI評審")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # API Key configuration section
        api_layout = QHBoxLayout()
        api_layout.addWidget(QLabel("Gemini API Key:"))

        # API Key status label
        self.api_key_status = QLabel("未設定")
        self.api_key_status.setStyleSheet("color: gray;")
        api_layout.addWidget(self.api_key_status)

        # API Key manage button
        manage_key_btn = QPushButton("API Key 設定")
        manage_key_btn.clicked.connect(self.manage_api_key)
        api_layout.addWidget(manage_key_btn)

        api_layout.addStretch()
        layout.addLayout(api_layout)

        # Prompt selection section
        prompt_layout = QHBoxLayout()
        prompt_layout.addWidget(QLabel("使用プロンプト:"))

        # Prompt selection combo box
        self.prompt_combo = QComboBox()
        self.prompt_combo.setMinimumWidth(300)
        self.prompt_combo.currentTextChanged.connect(self.on_prompt_changed)
        prompt_layout.addWidget(self.prompt_combo)

        # Prompt management button
        prompt_manage_btn = QPushButton("管理")
        prompt_manage_btn.clicked.connect(self.manage_prompts)
        prompt_layout.addWidget(prompt_manage_btn)

        prompt_layout.addStretch()
        layout.addLayout(prompt_layout)

        # Model selection section
        model_layout = QHBoxLayout()
        model_layout.addWidget(QLabel("モデル選択:"))

        self.model_combo = QComboBox()
        self.model_combo.addItems([
            "gemini-2.5-pro",
            "gemini-2.0-flash-exp",
            "gemini-1.5-pro",
            "gemini-1.5-flash"
        ])
        self.model_combo.setCurrentText("gemini-2.5-pro")
        self.model_combo.currentTextChanged.connect(self.on_model_changed)
        model_layout.addWidget(self.model_combo)

        # Model description
        self.model_description = QLabel("最高品質（処理時間: 長い、503エラーの可能性: 高い）")
        self.model_description.setStyleSheet("color: gray; font-size: 10px;")
        model_layout.addWidget(self.model_description)

        model_layout.addStretch()
        layout.addLayout(model_layout)

        # PDF files display
        pdf_label_layout = QHBoxLayout()
        pdf_label_layout.addWidget(QLabel("対象PDFファイル:"))
        pdf_label_layout.addStretch()

        # Add PDF file selection buttons
        add_pdf_btn = QPushButton("PDFを追加")
        add_pdf_btn.clicked.connect(self.add_pdf_files)
        pdf_label_layout.addWidget(add_pdf_btn)

        delete_pdf_btn = QPushButton("選択を削除")
        delete_pdf_btn.clicked.connect(self.delete_selected_pdf)
        pdf_label_layout.addWidget(delete_pdf_btn)

        clear_pdf_btn = QPushButton("全てクリア")
        clear_pdf_btn.clicked.connect(self.clear_pdf_files)
        pdf_label_layout.addWidget(clear_pdf_btn)

        layout.addLayout(pdf_label_layout)

        # Use DragDropListWidget for drag & drop support (PDF and MD files)
        self.pdf_list = DragDropListWidget(parent=self, file_filter=[".pdf", ".md"])
        self.pdf_list.setMaximumHeight(120)
        self.pdf_list.setSelectionMode(QListWidget.SingleSelection)
        layout.addWidget(self.pdf_list)

        # Start review button
        self.review_btn = QPushButton("AI評審開始 ▶")
        self.review_btn.setMinimumHeight(40)
        self.review_btn.clicked.connect(self.start_api_review)
        layout.addWidget(self.review_btn)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Status label
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # Result preview
        layout.addWidget(QLabel("評審結果プレビュー:"))
        self.result_preview = QTextEdit()
        self.result_preview.setReadOnly(True)
        layout.addWidget(self.result_preview)

        # Navigation buttons
        nav_layout = QHBoxLayout()
        prev_btn = QPushButton("← 戻る")
        prev_btn.clicked.connect(self.controller.prev_step)
        nav_layout.addWidget(prev_btn)

        self.next_btn = QPushButton("次へ →")
        self.next_btn.clicked.connect(self.go_next)
        nav_layout.addWidget(self.next_btn, alignment=Qt.AlignRight)

        layout.addLayout(nav_layout)
        self.setLayout(layout)

    def load_api_key(self):
        """Load API key and model selection from config file"""
        try:
            config = ConfigManager.load_config()
            api_key = config['API']['gemini_api_key']
            if api_key and api_key != "YOUR_API_KEY_HERE":
                self.api_key = api_key
                self.update_api_key_status()

            # Load model selection
            model = config['API'].get('gemini_model', 'gemini-2.5-pro')
            self.model_combo.setCurrentText(model)
        except Exception as e:
            logger.warning(f"Could not load API key from config: {e}")
            self.api_key = ""
            self.update_api_key_status()

    def update_api_key_status(self):
        """Update API key status display"""
        if self.api_key and ConfigManager.validate_api_key(self.api_key):
            # Show masked key
            masked = self.api_key[:8] + "..." + self.api_key[-4:] if len(self.api_key) > 12 else "••••••"
            self.api_key_status.setText(f"設定済み ({masked})")
            self.api_key_status.setStyleSheet("color: green; font-weight: bold;")
        else:
            self.api_key_status.setText("未設定")
            self.api_key_status.setStyleSheet("color: gray;")

    def manage_api_key(self):
        """Open API key management dialog"""
        dialog = APIKeyDialog(self, self.api_key)

        if dialog.exec() == QDialog.Accepted:
            new_key = dialog.get_api_key()

            if new_key and ConfigManager.validate_api_key(new_key):
                self.api_key = new_key
                self.update_api_key_status()

                # Save to config
                try:
                    # Try to load existing config, or create default if doesn't exist
                    try:
                        config = ConfigManager.load_config()
                    except FileNotFoundError:
                        # Create default config structure
                        logger.info("Creating new config.ini file")
                        config = {
                            'API': {
                                'gemini_api_key': '',
                                'gemini_model': 'gemini-2.5-pro'
                            },
                            'Paths': {'default_output_dir': './output'},
                            'Settings': {
                                'temperature': '0',
                                'max_output_tokens': '8192',
                                'max_retries': '3'
                            }
                        }

                    config['API']['gemini_api_key'] = new_key
                    ConfigManager.save_config(config)
                    QMessageBox.information(self, "保存完了", "API Keyを保存しました")
                except Exception as e:
                    logger.error(f"Could not save API key: {e}")
                    QMessageBox.warning(self, "警告", f"API Keyの保存に失敗しました:\n{str(e)}")
            elif new_key:
                QMessageBox.warning(self, "エラー", "無効なAPI Keyです")
            else:
                # User cleared the key
                self.api_key = ""
                self.update_api_key_status()

    def on_model_changed(self, model_name):
        """Handle model selection change"""
        # Update description based on model
        descriptions = {
            "gemini-2.5-pro": "最高品質（処理時間: 長い、503エラーの可能性: 高い）",
            "gemini-2.0-flash-exp": "高速・実験版（処理時間: 短い、503エラーの可能性: 低い）",
            "gemini-1.5-pro": "バランス型（処理時間: 中、503エラーの可能性: 中）",
            "gemini-1.5-flash": "最速（処理時間: 最短、503エラーの可能性: 最低）"
        }
        self.model_description.setText(descriptions.get(model_name, ""))

        # Save to config
        try:
            try:
                config = ConfigManager.load_config()
            except FileNotFoundError:
                # Create default config structure
                logger.info("Creating new config.ini file")
                config = {
                    'API': {
                        'gemini_api_key': self.api_key or '',
                        'gemini_model': 'gemini-2.5-pro'
                    },
                    'Paths': {'default_output_dir': './output'},
                    'Settings': {
                        'temperature': '0',
                        'max_output_tokens': '8192',
                        'max_retries': '3'
                    }
                }

            config['API']['gemini_model'] = model_name
            ConfigManager.save_config(config)
            logger.info(f"Model changed to: {model_name}")
        except Exception as e:
            logger.warning(f"Could not save model selection: {e}")

    def set_pdf_files(self, pdf_files):
        """Receive PDF file list from Step 1"""
        self.pdf_files = pdf_files
        self.update_pdf_list()

    def add_pdf_files(self):
        """Add PDF or Markdown files manually"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "PDFまたはMarkdownファイルを選択",
            "./output/pdfs",
            "Supported Files (*.pdf *.md);;PDF Files (*.pdf);;Markdown Files (*.md);;All Files (*.*)"
        )
        if file_paths:
            # Add to existing list (avoid duplicates)
            added_count = 0
            for path in file_paths:
                if path not in self.pdf_files:
                    self.pdf_files.append(path)
                    added_count += 1

            self.update_pdf_list()
            logger.info(f"Added {added_count} files manually")

    def add_dropped_files(self, file_paths):
        """Handle files dropped on PDF list"""
        added_count = 0
        for path in file_paths:
            if path not in self.pdf_files:
                self.pdf_files.append(path)
                added_count += 1

        if added_count > 0:
            self.update_pdf_list()
            logger.info(f"Added {added_count} PDF files via drag & drop")

    def clear_pdf_files(self):
        """Clear all PDF files"""
        reply = QMessageBox.question(
            self,
            "確認",
            "すべてのPDFファイルをクリアしますか？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.pdf_files.clear()
            self.update_pdf_list()
            logger.info("Cleared all PDF files")

    def update_pdf_list(self):
        """Update PDF list display"""
        self.pdf_list.clear()
        if self.pdf_files:
            for pdf_path in self.pdf_files:
                item = QListWidgetItem(Path(pdf_path).name)
                item.setToolTip(pdf_path)  # Show full path on hover
                self.pdf_list.addItem(item)
        else:
            item = QListWidgetItem("（PDFファイルがありません）")
            item.setFlags(item.flags() & ~Qt.ItemIsSelectable)  # Make it non-selectable
            self.pdf_list.addItem(item)

    def delete_selected_pdf(self):
        """Delete selected PDF from list"""
        current_row = self.pdf_list.currentRow()
        if current_row >= 0 and current_row < len(self.pdf_files):
            removed_file = self.pdf_files.pop(current_row)
            self.update_pdf_list()
            logger.info(f"Removed PDF file: {Path(removed_file).name}")
        else:
            QMessageBox.information(self, "情報", "削除するPDFファイルを選択してください")

    def load_prompts(self):
        """Load available prompts into combo box"""
        prompts_dir = Path("./prompts")
        prompts_dir.mkdir(exist_ok=True)

        # Save current selection
        current_text = self.prompt_combo.currentText()

        # Clear and reload
        self.prompt_combo.clear()

        # Get all prompt files
        prompt_files = sorted(prompts_dir.glob("*.txt"))

        if not prompt_files:
            # Create default if none exists
            default_prompt = """# AI評審プロンプト

このファイルにプロンプトを記入してください。"""
            default_file = prompts_dir / "標準テンプレート.txt"
            with open(default_file, 'w', encoding='utf-8') as f:
                f.write(default_prompt)
            prompt_files = [default_file]

        # Add to combo box
        for prompt_file in prompt_files:
            self.prompt_combo.addItem(prompt_file.stem)

        # Restore selection or select first
        if current_text:
            index = self.prompt_combo.findText(current_text)
            if index >= 0:
                self.prompt_combo.setCurrentIndex(index)
            else:
                self.prompt_combo.setCurrentIndex(0)
        else:
            self.prompt_combo.setCurrentIndex(0)

        # Update current path
        self.update_current_prompt_path()

    def on_prompt_changed(self, prompt_name):
        """Handle prompt selection change"""
        if prompt_name:
            prompts_dir = Path("./prompts")
            prompt_file = prompts_dir / f"{prompt_name}.txt"

            if prompt_file.exists():
                self.current_prompt_path = str(prompt_file)
                logger.info(f"Prompt changed to: {prompt_name}")
            else:
                logger.warning(f"Prompt file not found: {prompt_file}")

    def update_current_prompt_path(self):
        """Update current prompt path from combo box selection"""
        prompt_name = self.prompt_combo.currentText()
        if prompt_name:
            prompts_dir = Path("./prompts")
            prompt_file = prompts_dir / f"{prompt_name}.txt"
            if prompt_file.exists():
                self.current_prompt_path = str(prompt_file)

    def manage_prompts(self):
        """Open prompt template manager"""
        dialog = PromptManagerDialog(self)
        dialog.exec()

        # Always reload prompts after closing manager
        self.load_prompts()
        logger.info("Prompt list reloaded after management")

    def start_api_review(self):
        """Start API-based AI review"""
        if not self.pdf_files:
            QMessageBox.warning(self, "エラー", "PDFファイルがありません\n\n第1段階でPDFを生成するか、「PDFを追加」ボタンでPDFファイルを選択してください")
            return

        # Validate API key
        if not self.api_key or not ConfigManager.validate_api_key(self.api_key):
            QMessageBox.critical(
                self,
                "API Key エラー",
                "API Keyが設定されていません。\n\n"
                "「API Key 設定」ボタンをクリックして\n"
                "Gemini API Keyを設定してください。"
            )
            return

        try:
            # Get selected model from UI
            model = self.model_combo.currentText()

            # Load prompt from current selection
            if not self.current_prompt_path:
                raise FileNotFoundError("プロンプトが選択されていません\n\n使用プロンプトを選択してください")

            if not Path(self.current_prompt_path).exists():
                raise FileNotFoundError(f"プロンプトファイルが見つかりません:\n{self.current_prompt_path}")

            with open(self.current_prompt_path, 'r', encoding='utf-8') as f:
                prompt = f.read()

            prompt_name = self.prompt_combo.currentText()
            logger.info(f"Using prompt: {prompt_name}")

            # Disable button and show progress
            self.review_btn.setEnabled(False)
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)  # Indeterminate mode
            self.status_label.setText("AI評審を実行中...")

            # Create worker thread
            self.worker = AIReviewWorker(self.pdf_files, prompt, self.api_key, model)
            self.worker.finished.connect(self.on_review_finished)
            self.worker.error.connect(self.on_review_error)
            self.worker.progress.connect(self.on_review_progress)
            self.worker.start()

        except Exception as e:
            self.reset_ui()
            QMessageBox.critical(self, "エラー", f"AI評審開始失敗:\n{str(e)}")

    def on_review_progress(self, message):
        """Update progress status"""
        self.status_label.setText(message)

    def on_review_finished(self, result):
        """Handle successful review completion"""
        self.reset_ui()
        self.review_result = result

        # Display result preview
        preview = f"[完了] 評審完了\n\n抽出された行数: {len(result.rows)}\n\n"
        preview += "最初の3行のプレビュー:\n" + "="*50 + "\n\n"

        for i, row in enumerate(result.rows[:3], 1):
            preview += f"【{i}】 {row.requirement_no}\n"
            preview += f"  要求内容: {row.requirement_content[:50]}...\n"
            preview += f"  評価: {row.evaluation}\n"
            preview += f"  対応有無: {row.response_status}\n\n"

        self.result_preview.setText(preview)
        self.status_label.setText("[完了] AI評審が完了しました")
        self.status_label.setStyleSheet("color: green; font-weight: bold;")

        # Pass to Step 3
        self.controller.step3_page.set_review_result(result)

        QMessageBox.information(
            self,
            "完了",
            f"AI評審が完了しました\n\n抽出行数: {len(result.rows)}"
        )

    def on_review_error(self, error_msg):
        """Handle review error"""
        self.reset_ui()
        self.status_label.setText("[失敗] AI評審が失敗しました")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        QMessageBox.critical(self, "エラー", f"AI評審失敗:\n\n{error_msg}")

    def reset_ui(self):
        """Reset UI elements"""
        self.review_btn.setEnabled(True)
        self.progress.setVisible(False)

    def go_next(self):
        """Go to next step"""
        if not self.review_result:
            QMessageBox.warning(
                self,
                "警告",
                "AI評審が完了していません\n\n先にAI評審を実行してください"
            )
            return
        self.controller.next_step()


class Step3Page(QWidget):
    """Step 3: Save results to Excel"""
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.review_result = None

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Title
        title = QLabel("第3段階: 結果保存")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Form
        form = QFormLayout()

        # Version number input (from Step1 or manual input)
        self.version_input = QLineEdit()
        self.version_input.setPlaceholderText("例: 6")
        form.addRow("バージョン番号:", self.version_input)

        # Output directory selection
        dir_layout = QHBoxLayout()
        self.dir_label = QLabel("./output/results")
        dir_layout.addWidget(self.dir_label, stretch=1)

        browse_dir_btn = QPushButton("変更")
        browse_dir_btn.clicked.connect(self.browse_directory)
        dir_layout.addWidget(browse_dir_btn)
        form.addRow("出力先:", dir_layout)

        layout.addLayout(form)

        # Save button
        save_btn = QPushButton("Excelに保存 ▶")
        save_btn.setMinimumHeight(40)
        save_btn.clicked.connect(self.save_excel)
        layout.addWidget(save_btn)

        # Status label
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # Spacer
        layout.addStretch()

        # Navigation buttons
        nav_layout = QHBoxLayout()

        prev_btn = QPushButton("← 戻る")
        prev_btn.clicked.connect(self.controller.prev_step)
        nav_layout.addWidget(prev_btn)

        restart_btn = QPushButton("最初に戻る")
        restart_btn.clicked.connect(self.restart_from_step1)
        restart_btn.setStyleSheet("color: blue;")
        nav_layout.addWidget(restart_btn)

        layout.addLayout(nav_layout)

        self.setLayout(layout)

    def set_review_result(self, review_result):
        """Receive review result from Step 2"""
        self.review_result = review_result

    def update_version_number(self):
        """Update version number from controller"""
        if self.controller.version_number is not None:
            # Fill the text box with version number from Step1
            self.version_input.setText(str(self.controller.version_number))
            logger.info(f"Step3 auto-filled version number: {self.controller.version_number}")
        else:
            # Leave empty for user to input manually
            self.version_input.clear()
            logger.info("Step3 version number not set, user can input manually")

    def restart_from_step1(self):
        """Go back to Step 1"""
        reply = QMessageBox.question(
            self,
            "確認",
            "最初のステップに戻りますか？\n\n現在の作業状態は保持されます。",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.controller.go_to_step(0)  # Go to Step 1 (index 0)

    def browse_directory(self):
        """Select output directory"""
        dir_path = QFileDialog.getExistingDirectory(self, "出力先を選択")
        if dir_path:
            self.dir_label.setText(dir_path)

    def save_excel(self):
        """Save to Excel"""
        if not self.review_result:
            QMessageBox.warning(
                self,
                "エラー",
                "評審結果がありません\n\n第2段階でAI評審を実行してください"
            )
            return

        # Get version number from text input
        version_text = self.version_input.text().strip()
        if not version_text:
            QMessageBox.warning(
                self,
                "エラー",
                "バージョン番号を入力してください"
            )
            return

        if not version_text.isdigit():
            QMessageBox.warning(
                self,
                "エラー",
                "バージョン番号は数字で入力してください"
            )
            return

        version_number = int(version_text)

        try:
            output_path = step3.save_to_excel(
                self.review_result,
                version_number,
                self.dir_label.text()
            )

            self.status_label.setText(f"[成功] 保存成功: {Path(output_path).name}")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")

            logger.info(f"Excel saved to: {output_path}")

            # Automatically open the Excel file
            self.open_excel_file(output_path)

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存失敗:\n\n{str(e)}")
            logger.error(f"Excel save error: {str(e)}")

    def open_excel_file(self, file_path):
        """Open Excel file in default application"""
        import os
        import subprocess

        try:
            # Use os.startfile on Windows to open with default application
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])

            logger.info(f"Opened Excel file: {file_path}")
            QMessageBox.information(
                self,
                "完了",
                f"Excelファイルを保存して開きました:\n\n{Path(file_path).name}"
            )
        except Exception as e:
            logger.error(f"Failed to open Excel file {file_path}: {str(e)}")
            QMessageBox.warning(
                self,
                "警告",
                f"Excelファイルは保存されましたが、自動で開けませんでした:\n\n{file_path}\n\nエラー: {str(e)}"
            )


class ReviAIApp(QMainWindow):
    """Main application window"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ReviAI - AI評審作業システム")
        self.setMinimumSize(900, 700)

        # Store version number from Step1
        self.version_number = None

        # Create stacked widget for pages
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # Create pages
        self.step1_page = Step1Page(self)
        self.step2_page = Step2Page(self)
        self.step3_page = Step3Page(self)

        # Add pages to stack
        self.stack.addWidget(self.step1_page)
        self.stack.addWidget(self.step2_page)
        self.stack.addWidget(self.step3_page)

        logger.info("ReviAI application started")

    def next_step(self):
        """Go to next step"""
        current = self.stack.currentIndex()
        if current < self.stack.count() - 1:
            self.stack.setCurrentIndex(current + 1)
            logger.info(f"Navigated to step {current + 2}")

            # Update Step3 version number when navigating to Step3
            if current + 1 == 2:  # Step3 (index 2)
                self.step3_page.update_version_number()

    def prev_step(self):
        """Go to previous step"""
        current = self.stack.currentIndex()
        if current > 0:
            self.stack.setCurrentIndex(current - 1)
            logger.info(f"Navigated to step {current}")

    def go_to_step(self, step_index):
        """Go to specific step"""
        if 0 <= step_index < self.stack.count():
            self.stack.setCurrentIndex(step_index)
            logger.info(f"Navigated to step {step_index + 1}")


def main():
    """Main entry point"""
    app = QApplication(sys.argv)

    # Set application style (Fusion for modern look)
    app.setStyle("Fusion")

    # Create and show main window
    window = ReviAIApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
