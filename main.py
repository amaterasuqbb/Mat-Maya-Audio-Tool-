import sys
import os
import webbrowser
import wave
import tempfile
import shutil
from pydub import AudioSegment
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QTreeWidgetItem, QDialog, QMessageBox, QAbstractItemView, QMessageBox, QMenu
from PyQt6.QtCore import QFileInfo, QObject, QThread, pyqtSignal, Qt, QMimeData
from PyQt6.QtGui import QAction
from mat import Ui_MainWindow
from mat_about import Ui_About_Dialog
from mat_progressbar import Ui_Dialog

# Load HTML content from files
HTML_DIR = os.path.join(os.path.dirname(__file__), 'assets')

def load_html_file(self, filename):
    file_path = os.path.join(HTML_DIR, filename)
    if not os.path.exists(file_path):
        QMessageBox.warning(self, "Error", f"HTML file not found: {filename}")
        return ""  # Return empty string if not found
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()  # Return the file content

class AboutDialog(QDialog, Ui_About_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)

        # Make the "Get involved" label a clickable link.
        self.label_5.setText('<a href="https://github.com/your-project-link" style="color: rgb(0, 85, 255);">Get involved</a>')
        self.label_5.setOpenExternalLinks(True)

        self.pushButton_about.clicked.connect(self.show_about_text)
        self.pushButton_author.clicked.connect(self.show_author_text)
        self.pushButton_license.clicked.connect(self.show_license_text)

        # Set the initial text for the browser.
        self.show_about_text()

    def show_about_text(self):
        html_content = load_html_file(self, "about.html")
        self.textBrowser.setHtml(html_content)

    def show_author_text(self):
        html_content = load_html_file(self, "authors.html")
        self.textBrowser.setHtml(html_content)

    def show_license_text(self):
        html_content = load_html_file(self, "license.html")
        self.textBrowser.setHtml(html_content)

class ProgressDialog(QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowTitle("Converting Files...")
        self.setModal(True)

class ConvertWorker(QObject):
    # Signals to communicate with the main UI thread
    progress_updated = pyqtSignal(int)
    finished = pyqtSignal()
    file_converted = pyqtSignal(object) # object is a QTreeWidgetItem

    def __init__(self, items, temp_dir):
        super().__init__()
        self.items = items
        self.temp_dir = temp_dir

    def run(self):
        converted_count = 0
        total_items = len(self.items)

        for item in self.items:
            try:
                is_converted = item.data(1, Qt.ItemDataRole.UserRole)
                if is_converted:
                    converted_count += 1
                    self.progress_updated.emit(converted_count)
                    continue

                temp_file_path = item.data(0, Qt.ItemDataRole.UserRole)
                audio = AudioSegment.from_file(temp_file_path)
                converted_audio = audio.set_frame_rate(44100).set_sample_width(2).set_channels(2)
                output_file_name = os.path.splitext(os.path.basename(temp_file_path))[0] + ".wav"
                output_path = os.path.join(self.temp_dir, output_file_name)
                converted_audio.export(output_path, format="wav")

                os.remove(temp_file_path)
                item.setData(0, Qt.ItemDataRole.UserRole, output_path)
                item.setData(1, Qt.ItemDataRole.UserRole, True)

                # Re-check the WAV properties of the newly converted file
                support_maya_status = "No"
                try:
                    with wave.open(output_path, 'r') as w:
                        nchannels = w.getnchannels()
                        sampwidth = w.getsampwidth()
                        framerate = w.getframerate()
                        if framerate == 44100 and (sampwidth * 8) == 16 and nchannels == 2:
                            support_maya_status = "Yes"
                except (wave.Error, FileNotFoundError):
                    support_maya_status = "N/A"

                # Update the TreeWidget item with the new converted file's info
                item.setText(1, os.path.basename(output_path))
                item.setText(2, datetime.fromtimestamp(os.path.getmtime(output_path)).strftime('%Y-%m-%d %H:%M:%S'))
                item.setText(3, "wav")
                new_file_size = os.path.getsize(output_path)
                item.setText(4, f"{new_file_size / (1024 * 1024):.2f} MB")
                item.setText(5, support_maya_status)
                item.setText(6, "Complete")
                item.setText(7, "OK")

                converted_count += 1
                self.progress_updated.emit(converted_count)

            except Exception as e:
                print(f"Failed to convert: {e}")
                item.setText(6, "Error")
                item.setText(7, "Failed")
                converted_count += 1
                self.progress_updated.emit(converted_count)

        self.finished.emit()

class MatMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.showMaximized()

        self.treeWidget.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        self.treeWidget.setAcceptDrops(True)

        self.treeWidget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.treeWidget.customContextMenuRequested.connect(self.show_context_menu)

        # Create a temporary directory to store files
        self.temp_dir = tempfile.mkdtemp()
        self.connect_signals()
        self.treeWidget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def __del__(self):
        # Clean up the temporary directory when the application closes
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def connect_signals(self):
        # Connect buttons to functions.
        self.pushButton_1.clicked.connect(self.add_files)
        self.pushButton_2.clicked.connect(self.delete_selection)
        self.pushButton_3.clicked.connect(self.clear_list)
        self.pushButton_4.clicked.connect(self.show_file)
        self.pushButton_5.clicked.connect(self.convert_selection)
        self.pushButton_6.clicked.connect(self.convert_all)
        self.pushButton_7.clicked.connect(self.browse_folder)
        self.pushButton_8.clicked.connect(self.download_files)

        # Connect menu actions to functions.
        self.actionAdd_files.triggered.connect(self.add_files)
        self.actionExit.triggered.connect(self.exit_application)
        self.action_Delete.triggered.connect(self.delete_selection)
        self.actionSelect_All.triggered.connect(self.select_all_action)
        self.actionClear.triggered.connect(self.clear_list)
        self.actionConvert_Selection.triggered.connect(self.convert_selection)
        self.actionConvert_All.triggered.connect(self.convert_all)
        self.actionDownload.triggered.connect(self.download_files)
        self.actionHelp_Portal.triggered.connect(self.open_help_portal)
        self.actionVisit_Website.triggered.connect(self.visit_website)
        self.actionJoin_in_Discord_Server.triggered.connect(self.join_discord_server)
        self.actionAbout.triggered.connect(self.about_mat)

    def dragEnterEvent(self, event):
        # This method is called when a drag operation enters the widget
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        # This method is called when a drop is performed
        urls = event.mimeData().urls()
        for url in urls:
            # Check if the URL is a local file
            if url.isLocalFile():
                file_path = url.toLocalFile()
                self.add_file_to_treewidget(file_path)
        event.acceptProposedAction()

    def show_context_menu(self, pos):
        # Create the menu
        menu = QMenu(self)
    
        # Create actions for the menu
        add_action = QAction("Add Files", self)
        delete_action = QAction("Delete Selection", self)
        clear_action = QAction("Clear", self)
        convert_selection_action = QAction("Convert Selection", self)
        convert_all_action = QAction("Convert All", self)
        download_action = QAction("Download", self)
    
        # Connect actions to your existing functions
        add_action.triggered.connect(self.add_files)
        delete_action.triggered.connect(self.delete_selection)
        clear_action.triggered.connect(self.clear_list)
        convert_selection_action.triggered.connect(self.convert_selection)
        convert_all_action.triggered.connect(self.convert_all)
        download_action.triggered.connect(self.download_files)
    
        # Add actions to the menu
        menu.addAction(add_action)
        menu.addAction(delete_action)
        menu.addAction(clear_action)
        menu.addSeparator() # Adds a line to separate groups
        menu.addAction(convert_selection_action)
        menu.addAction(convert_all_action)
        menu.addSeparator()
        menu.addAction(download_action)
    
        # Show the menu at the cursor's position
        menu.exec(self.treeWidget.mapToGlobal(pos))

    # add files with button.
    def add_files(self):
        add_files_filter = (
            "All Media Files (*.mp4 *.avi *.mkv *.mov *.wmv *.flv *.webm *.mp3 *.wav *.flac *.aac *.m4a *.ogg *.aiff);;"
            "Video Files (*.mp4 *.avi *.mkv *.mov *.wmv *.flv *.webm);;"
            "Audio Files (*.mp3 *.wav *.flac *.aac *.m4a *.ogg *.aiff)"
        )
        filenames, _ = QFileDialog.getOpenFileNames(
            self,
            "Select one or more files to open",
            filter=add_files_filter
        )
        for file_path in filenames:
            self.add_file_to_treewidget(file_path)

    def add_file_to_treewidget(self, file_path):
        # Check for duplicate file names
        file_name = os.path.basename(file_path)
        for i in range(self.treeWidget.topLevelItemCount()):
            item = self.treeWidget.topLevelItem(i)
            if item.text(1) == file_name:
                QMessageBox.warning(self, "Duplicate File", f"The file '{file_name}' already exists in the list.")
                return

        # Copy the file to the temporary directory
        temp_file_path = os.path.join(self.temp_dir, os.path.basename(file_path))
        shutil.copy(file_path, temp_file_path)

        # Initialize all variables with default values
        file_name = "N/A"
        file_date_modified = "N/A"
        file_type = "N/A"
        file_size = "N/A"
        support_maya_status = "No"
        convert_progress = "N/A"
        convert_status = "N/A"

        # Get the number of items already in the tree.
        current_item_count = self.treeWidget.topLevelItemCount()
        new_item_number = current_item_count + 1

        try:
            # Get the file information. This section is robustly handled by a try-except block.
            file_name = os.path.basename(file_path)
            file_type = os.path.splitext(file_path)[1][1:].lower()

            # Get modification time
            mod_timestamp = os.path.getmtime(file_path)
            file_date_modified = datetime.fromtimestamp(mod_timestamp).strftime('%Y-%m-%d %H:%M %p')

            # Get file size
            file_size_bytes = os.path.getsize(file_path)
            file_size = f"{file_size_bytes / (1024 * 1024):.2f} MB"

            # Check if the file is a WAV and meets specific criteria
            if file_type == 'wav':
                with wave.open(file_path, 'r') as w:
                    nchannels = w.getnchannels()
                    sampwidth = w.getsampwidth()
                    framerate = w.getframerate()

                    # Check for 44.1kHz, 16bit, 2 channels, 1411kbps.
                    if framerate == 44100 and (sampwidth * 8) == 16 and nchannels == 2:
                        support_maya_status = "Yes"

        except (OSError, FileNotFoundError, wave.Error):
            # This block catches any error during file access or parsing.
            # Variables will remain as their default "N/A" or "No" values.
            pass

        item = QTreeWidgetItem(self.treeWidget)
        item.setText(0, str(new_item_number))
        item.setText(1, file_name)
        item.setText(2, file_date_modified)
        item.setText(3, file_type)
        item.setText(4, file_size)
        item.setText(5, support_maya_status)
        item.setText(6, convert_progress)
        item.setText(7, convert_status)

        item.setData(0, Qt.ItemDataRole.UserRole, temp_file_path)
        item.setData(1, Qt.ItemDataRole.UserRole, False)

        # Add the item to the tree widget
        self.treeWidget.addTopLevelItem(item)
        
        # Force an update of the widget to ensure it's displayed
        self.treeWidget.repaint()

    def open_file_dialog(self):
        file_paths, _ = QFileDialog.getOpenFileName(self, "Select one or more files to open")
        if file_path:
                for file_path in file_paths:
                    self.add_file_to_treewidget(file_path)

    def delete_selection(self):
        selected_items = self.treeWidget.selectedItems()

        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select one or more items to delete.")
            return

        # Delete the selected items and their temporary files
        for item in selected_items:
            parent = item.parent()
            temp_file_path = item.data(0, Qt.ItemDataRole.UserRole)

            # Remove the temporary file from the disk
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)

            if parent:
                parent.removeChild(item)
            else:
                index = self.treeWidget.indexOfTopLevelItem(item)
                self.treeWidget.takeTopLevelItem(index)

        # After deletion, re-number all the remaining items
        for i in range(self.treeWidget.topLevelItemCount()):
            item = self.treeWidget.topLevelItem(i)
            item.setText(0, str(i + 1)) # Re-assign the number

        QMessageBox.information(self, "Deletion Complete", f"{len(selected_items)} item(s) have been deleted.")
        self.treeWidget.repaint()

    def clear_list(self):
        # This will remove all items from the tree widget
        self.treeWidget.clear()

        # Optional: You can show a message box to confirm the action
        QMessageBox.information(self, "List Cleared", "All items have been removed from the list.")

    def show_file(self):
        # Get the search query from the line edit and convert to lowercase for case-insensitive search
        search_text = self.lineEdit_1.text().strip().lower()

        if not search_text:
            QMessageBox.information(self, "Search", "Please enter a file name to search for.")
            return
    
        # Clear any previous selections
        self.treeWidget.clearSelection()

        # Loop through all top-level items in the tree widget
        for i in range(self.treeWidget.topLevelItemCount()):
            item = self.treeWidget.topLevelItem(i)

            # Get the file name from column 1 (assuming it's the second column)
            # Convert it to lowercase for comparison
            file_name = item.text(1).lower()

            if search_text in file_name:
                # If a match is found, select the item
                item.setSelected(True)

                # Scroll to the item to make it visible
                self.treeWidget.scrollToItem(item)

                # Stop searching after the first match is found
                return
        
        # If the loop finishes without finding a match, show a message box
        QMessageBox.information(self, "Search Results", f"No file found with the name '{search_text}'.")

    def convert_selection(self):
        selected_items = self.treeWidget.selectedItems()

        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select one or more items to convert.")
            return

        # Set up and show the progress dialog
        self.progress_dialog = ProgressDialog(self)
        self.progress_dialog.label_6.setText("0")
        self.progress_dialog.label_5.setText(str(len(selected_items)))
        self.progress_dialog.progressBar.setMaximum(len(selected_items))
        self.progress_dialog.show()

        # Create the thread and worker
        self.thread = QThread()
        self.worker = ConvertWorker(selected_items, self.temp_dir)
        self.worker.moveToThread(self.thread)

        # Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.progress_updated.connect(self.update_progress_bar)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def update_progress_bar(self, count):
        # This slot updates the progress dialog
        self.progress_dialog.label_6.setText(str(count))
        self.progress_dialog.progressBar.setValue(count)

    def on_conversion_finished(self):
        self.progress_dialog.close()
        QMessageBox.information(self, "Conversion Complete", "Selected files have been converted successfully!")
        self.treeWidget.repaint()

    def convert_all(self):
        reply = QMessageBox.question(self, "Convert All", "Are you sure you want to convert all files?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    
        if reply == QMessageBox.StandardButton.No:
            return
            
        total_items = self.treeWidget.topLevelItemCount()
        all_items = [self.treeWidget.topLevelItem(i) for i in range(total_items)]
    
        # Set up and show the progress dialog
        self.progress_dialog = ProgressDialog(self)
        self.progress_dialog.label_6.setText("0")
        self.progress_dialog.label_5.setText(str(total_items))
        self.progress_dialog.progressBar.setMaximum(total_items)
        self.progress_dialog.show()
    
        # Create the thread and worker
        self.thread = QThread()
        self.worker = ConvertWorker(all_items, self.temp_dir)
        self.worker.moveToThread(self.thread)
    
        # Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.progress_updated.connect(self.update_progress_bar)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
    
        self.thread.start()

    def browse_folder(self):
        # Open the file explorer to select a directory
        selected_folder = QFileDialog.getExistingDirectory(self, "Select Download Folder")

        if selected_folder:
            # Check if the folder is already in the QComboBox
            index = self.comboBox.findText(selected_folder)

            if index == -1:
                # If not in the list, add it to the top
                self.comboBox.insertItem(0, selected_folder)
                self.comboBox.setCurrentIndex(0)
            else:
                # If it's already in the list, just set it as the current selection
                self.comboBox.setCurrentIndex(index)

        # In your export/download function
        download_path = self.comboBox.currentText()
        if not download_path:
            QMessageBox.warning(self, "No Folder Selected", "Please select a download folder first.")
            return

    def update_download_path(self):
        current_path = self.comboBox.currentText()
        
        # Check if the path is a valid directory
        if os.path.isdir(current_path):
            self.download_path = current_path
            print(f"Download path set to: {self.download_path}")
        else:
            # Optionally, warn the user if the path is not valid
            # QMessageBox.warning(self, "Invalid Path", "The path you entered is not a valid directory.")
            self.download_path = None
            
        # Optional: If you want to automatically add the typed path to the list
        if os.path.isdir(current_path) and self.comboBox.findText(current_path) == -1:
            self.comboBox.insertItem(0, current_path)

    def download_files(self):
        selected_items = self.treeWidget.selectedItems()
        download_path = self.comboBox.currentText()

        # Step 1: Check for a valid download path
        if not download_path or not os.path.isdir(download_path):
            QMessageBox.warning(self, "Invalid Path", "Please select a valid download folder first.")
            return

        # Step 2: Check if any items are selected
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select one or more items to download.")
            return
    
        # Step 3: Loop through and check each file's conversion status
        for item in selected_items:
            is_converted = item.data(1, Qt.ItemDataRole.UserRole)

            if not is_converted:
                QMessageBox.warning(self, "File Not Converted", f"The file '{item.text(1)}' has not been converted. Please convert it first.")
                return # Exit the function if any selected file is not converted

        # Step 4: If all checks pass, start the download process
        for item in selected_items:
            try:
                temp_file_path = item.data(0, Qt.ItemDataRole.UserRole)
                file_name = os.path.basename(temp_file_path)
                destination_path = os.path.join(download_path, file_name)

                shutil.copy(temp_file_path, destination_path)
                print(f"Downloaded {file_name} to {destination_path}")

            except Exception as e:
                QMessageBox.critical(self, "Download Error", f"Failed to download {file_name}.\nError: {e}")

        QMessageBox.information(self, "Download Complete", "All selected files have been downloaded successfully!")

    def exit_application(self):
        self.close()

    def delete_action(self):
        self.action_Delete.triggered.connect(self.delete_action)

    def select_all_action(self):
        self.treeWidget.selectAll()

    def open_help_portal(self):
        # change the URL after repository is set up
        webbrowser.open_new_tab("https://example.com/help")

    def visit_website(self):
        # change the URL after repository is set up
        webbrowser.open_new_tab("https://github.com/amaterasuqbb")

    def join_discord_server(self):
        # change the URL after repository is set up
        webbrowser.open_new_tab("https://discord.gg/6aTkgP6a")

    def about_mat(self):
        about_dialog = AboutDialog(self)
        about_dialog.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MatMainWindow()
    window.show()
    sys.exit(app.exec())