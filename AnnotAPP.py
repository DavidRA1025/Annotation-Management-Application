# AnnotAPP - Annotation Management Tool
# Copyright (C) 2024 chenwayi
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QTextEdit, 
                             QVBoxLayout, QHBoxLayout, QGridLayout, QWidget, 
                             QLabel, QCheckBox, QFileDialog, QMessageBox, QDialog,
                             QLineEdit, QInputDialog, QScrollArea, QFormLayout,
                             QDialogButtonBox, QAction, QMenu)
from PyQt5.QtCore import Qt, QSettings, QEvent, QObject, QRect, QPropertyAnimation, QEasingCurve, pyqtProperty, pyqtSignal
from PyQt5.QtGui import QFont, QClipboard, QTextCursor, QIcon, QPainter, QColor, QPen
import ctypes

VERSION = "3.0"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

# Use the resource_path function to get the correct path for the icon
icon_path = resource_path('AnnotAPP.ico')

# Set app ID for Windows 7 and above
myappid = 'amazon.cdt.annotapp.3.0'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

class ToggleSwitch(QWidget):
    stateChanged = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(40, 20)  # Smaller size, but not as small as previously suggested
        self._enabled = False
        self._margin = 2

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        
        if self._enabled:
            painter.setBrush(QColor(0, 255, 0))  # Green when enabled
        else:
            painter.setBrush(QColor(120, 120, 120))  # Gray when disabled
        
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 10, 10)
        
        painter.setBrush(QColor(255, 255, 255))
        if self._enabled:
            painter.drawEllipse(self.width() - self._margin - 16, self._margin, 16, 16)
        else:
            painter.drawEllipse(self._margin, self._margin, 16, 16)

    def mousePressEvent(self, event):
        self._enabled = not self._enabled
        self.update()
        self.stateChanged.emit(self._enabled)

    def isEnabled(self):
        return self._enabled

    def setEnabled(self, enabled):
        if self._enabled != enabled:
            self._enabled = enabled
            self.update()
            self.stateChanged.emit(self._enabled)

class SmartSelectTextEdit(QTextEdit):
    def mouseDoubleClickEvent(self, event):
        cursor = self.textCursor()
        pos = cursor.position()
        
        # Select the initial word
        cursor.select(QTextCursor.WordUnderCursor)
        
        # Get the full text and the initial selection
        full_text = self.toPlainText()
        start = cursor.selectionStart()
        end = cursor.selectionEnd()
        
        # Define characters that can be part of a compound word
        word_chars = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._-@')
        
        # Expand left
        while start > 0 and full_text[start-1] in word_chars:
            start -= 1
        
        # Expand right
        while end < len(full_text) and full_text[end] in word_chars:
            end += 1
        
        # Set the new selection
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.KeepAnchor)
        
        self.setTextCursor(cursor)

class VersionHistoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Version History")
        self.setMinimumSize(500, 400)  # Increased size for better readability
        
        layout = QVBoxLayout(self)
        
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        font = QFont("Consolas, Monaco, Monospace", 10)  # Use a monospace font
        self.text_edit.setFont(font)
        layout.addWidget(self.text_edit)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)
        
        self.set_version_history()

    def set_version_history(self):
        history = """
Version 3.0 (Current):
- Added new accumulative mode (toggle with Ctrl+M)
- Implemented stacking of annotations in accumulative mode
- Added right-click functionality to remove annotations
- Improved display with text wrapping and scrollbar
- Added Ctrl+L shortcut to clear display
- Fixed bug with toggle switch functionality
- Improved visual indication of active annotations in both modes
- Fixed highlighting of multiple active annotations in accumulative mode

Version 2.3:
- Removed "Edit All" button from main window
- Added "Edit All" functionality to Settings menu
- Improved button layout and window resizing
- Made version history window scrollable

Version 2.2:
- Added custom application icon with improved Windows integration
- Implemented real-time character count display in text area
- Retained "Contributor: chenwayi@" label and "Always on Top" functionality

Version 2.1:
- Implemented smart text selection for compound words and email addresses
- Double-clicking now selects entire email addresses, function names, etc.

Version 2.0:
- Renamed application to Annot APP
- Migrated from tkinter to PyQt5 for improved UI capabilities
- Implemented a grid layout for annotation buttons
- Added font size adjustment feature (Ctrl + mouse wheel)
- Set Consolas as the default font with fallbacks
- Added "Always on Top" functionality with a checkbox
- Enhanced error handling and user prompts
- Implemented plain text handling to remove formatting
- Added contributor information to the UI

Version 1.3:
- Introduced Edit mode for annotations (removed in 2.0)
- Added Edit and Save buttons to toggle and apply changes (removed in 2.0)
- Implemented temporary saving of edited annotations (removed in 2.0)

Version 1.2:
- Refined UI with better spacing and organization
- Added a contributor label
- Improved error handling and user feedback
- Enhanced "Remove All" functionality

Version 1.1:
- Improved UI layout with separate areas for display, buttons, and bottom information
- Added a Settings button for less frequent operations
- Implemented center window positioning on startup

Version 1.0 (Initial Release):
- Basic annotation management functionality (add, delete, display)
- Import and export annotations to/from Excel files
- Simple UI with a display area and buttons
        """
        self.text_edit.setPlainText(history)

class WheelEventFilter(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel and event.modifiers() == Qt.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                self.parent().change_font_size(True)
            else:
                self.parent().change_font_size(False)
            return True
        return False

class AnnotApp(QMainWindow):
    def __init__(self):  
        super().__init__()
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Annot APP {VERSION}")
        self.accumulative_mode = False  # Start in non-accumulative mode
        self.annotations = {}
        self.button_column = 0
        self.button_row = 0
        self.current_annotation = None
        self.annotation_file = "annotations.xlsx"
        self.settings = QSettings("MyCompany", "AnnotApp")
        self.active_annotations = set()

        self.display_text = SmartSelectTextEdit()
        self.wheel_event_filter = WheelEventFilter(self)
        self.display_text.viewport().installEventFilter(self.wheel_event_filter)
        
        # Main layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # Add toggle switch
        self.toggle_switch = ToggleSwitch()
        self.toggle_switch.stateChanged.connect(lambda state: self.toggle_mode(state))
        
        top_right_layout = QHBoxLayout()
        top_right_layout.addStretch()
        top_right_layout.addWidget(self.toggle_switch)
        main_layout.addLayout(top_right_layout)

        # UI layout
        self.setup_display_area(main_layout)
        self.setup_button_area(main_layout)
        self.setup_bottom_area(main_layout)

        self.setup_mode_toggle()
        self.setup_clear_shortcut()

        # Load annotations after UI setup
        self.load_annotations()

        # Set window size and position
        self.resize(600, 400)
        self.center_window()

        self.display_text.textChanged.connect(self.update_char_count)
        self.update_char_count()

    def setup_display_area(self, main_layout):
        display_layout = QHBoxLayout()

        # Configure display text
        self.display_text.setAcceptRichText(False)
        self.display_text.setLineWrapMode(QTextEdit.WidgetWidth)
        
        # Set Consolas font
        font = QFont("Consolas, Monaco, Monospace", 12)  # Fallback fonts included
        self.display_text.setFont(font)
        
        # Character count label
        self.char_count_label = QLabel("0")
        self.char_count_label.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.char_count_label.setStyleSheet("QLabel { background-color : white; padding: 1px; }")

        # Create a container for the text edit and the label
        container = QWidget()
        container_layout = QGridLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.addWidget(self.display_text, 0, 0)
        container_layout.addWidget(self.char_count_label, 0, 0, Qt.AlignRight | Qt.AlignTop)

        # Add the container to the display layout
        display_layout.addWidget(container)

        # Display button frame
        display_button_layout = QVBoxLayout()
        
        # Settings button
        settings_button = QPushButton("Settings")
        settings_button.clicked.connect(self.open_settings)
        display_button_layout.addWidget(settings_button)

        # Copy button
        copy_button = QPushButton("COPY")
        copy_button.clicked.connect(self.copy_annotation)
        display_button_layout.addWidget(copy_button)

        # Clear button
        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.clear_display)
        display_button_layout.addWidget(clear_button)

        # Add button layout to main display layout
        display_layout.addLayout(display_button_layout)

        # Add the entire display layout to the main layout
        main_layout.addLayout(display_layout)

    def update_char_count(self):
        count = len(self.display_text.toPlainText())
        self.char_count_label.setText(str(count))

    def setup_button_area(self, main_layout):
        self.button_layout = QGridLayout()
        main_layout.addLayout(self.button_layout)

    def setup_bottom_area(self, main_layout):
        bottom_layout = QHBoxLayout()

        # Contributor Label
        contributor_label = QLabel(f"Contributor: chenwayi@ | Version: {VERSION}")
        contributor_label.mousePressEvent = self.show_version_history
        bottom_layout.addWidget(contributor_label)

        # Always on Top Checkbox
        self.always_on_top_check = QCheckBox("Always on Top")
        self.always_on_top_check.stateChanged.connect(self.toggle_always_on_top)
        bottom_layout.addWidget(self.always_on_top_check)

        main_layout.addLayout(bottom_layout)

    def setup_mode_toggle(self):
        toggle_action = QAction('Toggle Mode', self)
        toggle_action.setShortcut('Ctrl+M')
        toggle_action.triggered.connect(lambda: self.toggle_mode())
        self.addAction(toggle_action)

    def toggle_mode(self, state=None):
        if state is None:
            self.accumulative_mode = not self.accumulative_mode
        else:
            self.accumulative_mode = state
        
        self.update_mode_indicator()
        
        # Update the toggle switch state
        self.toggle_switch.blockSignals(True)
        self.toggle_switch.setEnabled(self.accumulative_mode)
        self.toggle_switch.blockSignals(False)

    def update_mode_indicator(self):
        mode = "Accumulative" if self.accumulative_mode else "Non-accumulative"
        self.setWindowTitle(f"Annot APP {VERSION} - {mode} Mode")

    def setup_clear_shortcut(self):
        clear_action = QAction('Clear Display', self)
        clear_action.setShortcut('Ctrl+L')
        clear_action.triggered.connect(self.clear_display)
        self.addAction(clear_action)

    def clear_display(self):
        self.display_text.clear()
        self.active_annotations.clear()
        self.update_button_states()  # No argument needed here
        self.current_annotation = None

    def reset_button_states(self):
        for i in range(self.button_layout.count()):
            button = self.button_layout.itemAt(i).widget()
            button.setStyleSheet("")

    def update_buttons(self):
        # Clear existing buttons
        for i in reversed(range(self.button_layout.count())): 
            widget = self.button_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)

        # Add buttons in alphabetical order
        sorted_names = sorted(self.annotations.keys())
        for index, name in enumerate(sorted_names):
            row = index // 4  # 4 buttons per row
            col = index % 4
            btn = QPushButton(name)
            btn.clicked.connect(lambda checked, n=name: self.show_annotation(n))
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, n=name: self.on_context_menu(pos, n))
            self.button_layout.addWidget(btn, row, col)

        # Adjust window size
        self.adjustSize()

    def on_context_menu(self, pos, name):
        context_menu = QMenu(self)
        remove_action = context_menu.addAction("Remove")
        action = context_menu.exec_(self.sender().mapToGlobal(pos))
        if action == remove_action:
            self.remove_annotation(name)

    def remove_annotation(self, name):
        if name in self.annotations:
            annotation_text = self.annotations[name]
            del self.annotations[name]
            if self.accumulative_mode:
                current_text = self.display_text.toPlainText()
                if annotation_text in current_text:
                    new_text = current_text.replace(annotation_text, '', 1)
                    self.display_text.setPlainText(new_text.strip())
            self.active_annotations.discard(name)
            self.update_buttons()
            self.update_button_states()  # No argument needed here

    def show_annotation(self, name):
        if self.accumulative_mode:
            current_text = self.display_text.toPlainText()
            if current_text:
                current_text += '\n'
            self.display_text.setPlainText(current_text + self.annotations[name])
            self.active_annotations.add(name)
        else:
            self.display_text.setPlainText(self.annotations[name])
            self.active_annotations = {name}
        self.update_button_states()  # No argument needed here
        self.current_annotation = name

    def update_button_states(self):
        for i in range(self.button_layout.count()):
            button = self.button_layout.itemAt(i).widget()
            if button.text() in self.active_annotations:
                button.setStyleSheet("background-color: lightblue;")
            else:
                button.setStyleSheet("")

    def edit_all_annotations(self):
        if not self.annotations:
            QMessageBox.information(self, "No Annotations", "There are no annotations to edit.")
            return

        edit_dialog = QDialog(self)
        edit_dialog.setWindowTitle("Edit All Annotations")
        layout = QVBoxLayout(edit_dialog)

        # Create a form layout for name-content pairs
        form_layout = QFormLayout()
        text_edits = {}

        for name, content in self.annotations.items():
            label = QLabel(name)
            text_edit = QTextEdit()
            text_edit.setPlainText(content)
            form_layout.addRow(label, text_edit)
            text_edits[name] = text_edit

        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_widget.setLayout(form_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)

        # Add Save and Cancel buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        button_box.accepted.connect(edit_dialog.accept)
        button_box.rejected.connect(edit_dialog.reject)
        layout.addWidget(button_box)

        result = edit_dialog.exec_()
        if result == QDialog.Accepted:
            # Update annotations with new content
            for name, text_edit in text_edits.items():
                new_content = text_edit.toPlainText()
                if new_content != self.annotations[name]:
                    self.annotations[name] = new_content

            self.update_buttons()
            self.adjustSize()
            QMessageBox.information(self, "Success", "All annotations have been updated.")

    def open_settings(self):
        settings_dialog = QDialog(self)
        settings_dialog.setWindowTitle("Settings")
        settings_dialog.setFixedSize(180, 220)  # Increased height to accommodate the new button

        layout = QVBoxLayout(settings_dialog)

        buttons = [
            ("Add", self.add_annotation_button),
            ("Edit All", self.edit_all_annotations),
            ("Delete", self.delete_annotation),
            ("Import", self.import_annotations),
            ("Export", self.export_annotations),
            ("Remove All", self.remove_all_annotations)
        ]

        for text, command in buttons:
            button = QPushButton(text)
            button.clicked.connect(command)
            layout.addWidget(button)

        settings_dialog.exec_()

    def add_annotation_button(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Create Annotation")
        layout = QVBoxLayout(dialog)

        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Name:"))
        self.name_entry = QLineEdit()
        name_layout.addWidget(self.name_entry)
        layout.addLayout(name_layout)

        text_layout = QHBoxLayout()
        text_layout.addWidget(QLabel("Annotation:"))
        self.text_entry = SmartSelectTextEdit()
        self.text_entry.setAcceptRichText(False)
        text_layout.addWidget(self.text_entry)
        layout.addLayout(text_layout)

        save_button = QPushButton("Save")
        save_button.clicked.connect(lambda: self.save_annotation(dialog))
        layout.addWidget(save_button)

        dialog.exec_()

    def save_annotation(self, dialog, name=None, text=None):
        if name is None or text is None:
            name = self.name_entry.text().strip()
            text = self.text_entry.toPlainText().strip()
        if name and text:
            if name in self.annotations:
                QMessageBox.warning(self, "Duplicate Name", "An annotation with this name already exists. Please choose a different name.")
                return
            self.annotations[name] = text
            self.update_buttons()
            if dialog:
                dialog.accept()
        elif dialog:
            QMessageBox.warning(self, "Invalid Input", "Both name and annotation text must be provided.")

    def copy_annotation(self):
        annotation_text = self.display_text.toPlainText().strip()
        if annotation_text:
            QApplication.clipboard().setText(annotation_text, mode=QClipboard.Clipboard)

    def delete_annotation(self):
        if self.current_annotation:
            confirm = QMessageBox.question(self, "Confirm Deletion", 
                                           f"Are you sure you want to delete the annotation '{self.current_annotation}'?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if confirm == QMessageBox.Yes:
                del self.annotations[self.current_annotation]
                self.update_buttons()
                self.show_annotation(None)
                self.display_text.clear()
        else:
            QMessageBox.information(self, "No Selection", "Please select an annotation to delete.")

    def remove_all_annotations(self):
        confirm = QMessageBox.question(self, "Confirm Removal", "Are you sure you want to remove all annotations?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.annotations.clear()
            self.update_buttons()
            self.show_annotation(None)
            self.display_text.clear()

    def import_annotations(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Import Annotations", "", "Excel files (*.xlsx)")
        if filepath:
            try:
                df = pd.read_excel(filepath)
                if 'Name' not in df.columns or 'Annotation' not in df.columns:
                    raise ValueError("Excel file must have 'Name' and 'Annotation' columns")
                for index, row in df.iterrows():
                    name = row['Name']
                    annotation = row['Annotation']
                    if name in self.annotations:
                        overwrite = QMessageBox.question(self, "Overwrite?", f"Annotation '{name}' already exists. Overwrite?",
                                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                        if overwrite == QMessageBox.No:
                            continue
                    self.annotations[name] = annotation
                self.update_buttons()
                QMessageBox.information(self, "Success", "Annotations imported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def export_annotations(self):
        filepath, _ = QFileDialog.getSaveFileName(self, "Export Annotations", "", "Excel files (*.xlsx)")
        if filepath:
            try:
                df = pd.DataFrame(list(self.annotations.items()), columns=['Name', 'Annotation'])
                df.to_excel(filepath, index=False)
                QMessageBox.information(self, "Success", "Annotations exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export annotations: {e}")

    def load_annotations(self):
        if os.path.exists(self.annotation_file):
            try:
                df = pd.read_excel(self.annotation_file)
                for index, row in df.iterrows():
                    self.save_annotation(None, row['Name'], row['Annotation'])
            except pd.errors.EmptyDataError:
                pass  # Handle empty file

    def save_annotations_to_file(self):
        if self.annotations:
            df = pd.DataFrame(list(self.annotations.items()), columns=['Name', 'Annotation'])
            df.to_excel(self.annotation_file, index=False)

    def toggle_always_on_top(self, state):
        self.setWindowFlag(Qt.WindowStaysOnTopHint, state == Qt.Checked)
        self.show()

    def change_font_size(self, increase):
        current_font = self.display_text.font()
        size = current_font.pointSize()
        if increase:
            size = min(size + 1, 72)  # Max font size of 72
        else:
            size = max(size - 1, 6)   # Min font size of 6
        current_font.setPointSize(size)
        self.display_text.setFont(current_font)

    def center_window(self):
        qr = self.frameGeometry()
        cp = QApplication.desktop().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def closeEvent(self, event):
        self.save_annotations_to_file()
        event.accept()

    def show_version_history(self, event):
        dialog = VersionHistoryDialog(self)
        dialog.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set the app icon
    app_icon = QIcon(icon_path)
    app.setWindowIcon(app_icon)
    
    window = AnnotApp()
    window.show()
    sys.exit(app.exec_())