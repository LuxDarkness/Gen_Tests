'''This module defines the main window of the application.'''

import sys
import json
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QWidget, QFileDialog, QFrame, QCheckBox, QProgressBar, QMessageBox
from watcher import Watcher

class MainWindow(QMainWindow):
    '''Main window of the application.'''

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Consolidation")
        self.setGeometry(0, 0, 700, 250)
        self.center()

        self.sel_folder_text = None
        self.pross_folder_text = None
        self.not_app_folder_text = None
        self.file_text = None
        self.check_box = None
        self.exec_button = None
        self.state_label = None
        self.progress_bar = None
        self.active = False
        self.watcher = None

        self.define_ui()
        self.load_data()

    def center(self):
        '''Center the window on the screen.'''

        frame = self.frameGeometry()
        center = self.screen().availableGeometry().center()
        frame.moveCenter(center)
        self.move(frame.topLeft())

    def define_ui(self):
        '''Define the user interface of the main window.'''

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)

        main_layout = QVBoxLayout()

        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.Box)
        frame.setFrameShadow(QFrame.Shadow.Raised)
        frame.setLineWidth(1)

        frame_layout = QVBoxLayout()
        frame.setLayout(frame_layout)

        exec_frame = QFrame()
        exec_frame.setFrameShape(QFrame.Shape.Box)
        exec_frame.setFrameShadow(QFrame.Shadow.Raised)
        exec_frame.setLineWidth(1)

        exec_frame_layout = QVBoxLayout()
        exec_frame.setLayout(exec_frame_layout)

        sel_folder_layout = self.sel_folder_layout()
        pross_folder_layout = self.pross_folder_layout()
        not_app_folder_layout = self.not_app_folder_layout()
        file_layout = self.file_layout()

        frame_layout.addLayout(sel_folder_layout)
        frame_layout.addLayout(pross_folder_layout)
        frame_layout.addLayout(not_app_folder_layout)
        frame_layout.addLayout(file_layout)

        exec_layout = self.exec_layout()
        exec_frame_layout.addLayout(exec_layout)

        main_layout.addWidget(frame)
        lower_layout = QVBoxLayout()
        lower_layout.addWidget(exec_frame)
        main_layout.addLayout(lower_layout)
        main_widget.setLayout(main_layout)

    def load_data(self):
        '''Load the data from the JSON configuration file.'''

        with open("Resources/config.json", "r", encoding='UTF-8') as file:
            data = json.load(file)

        self.sel_folder_text.setText(data["Observe_Folder"])
        self.pross_folder_text.setText(data["Processed_Folder"])
        self.not_app_folder_text.setText(data["Not_Applicable_Folder"])
        self.file_text.setText(data["Consolidation_File"])

    def edit_json_field(self, field_name: str, new_value: str):
        '''Edit a given field in the JSON configuration file.

        Parameters
        ----------
        field_name : str
            The name of the field to be edited.
        new_value : str
            The new value to be assigned to the field.
        '''

        with open("Resources/config.json", "r", encoding='UTF-8') as file:
            data = json.load(file)

        data[field_name] = new_value

        with open("Resources/config.json", "w", encoding='UTF-8') as file:
            json.dump(data, file, indent=4)

    def sel_folder_layout(self) -> QHBoxLayout:
        '''Define the layout for selecting the folder to be observed.

        Returns
        -------
        sel_folder_layout : QHBoxLayout
            Layout for selecting the folder to be observed.
        '''

        sel_folder_layout = QHBoxLayout()
        sel_folder_label = QLabel("Folder to be observed:")
        sel_folder_label.setFixedWidth(170)
        self.sel_folder_text = QLineEdit()
        self.sel_folder_text.setFixedWidth(350)
        self.sel_folder_text.setReadOnly(True)
        sel_folder_button = QPushButton("Select Folder")
        sel_folder_button.setFixedWidth(100)

        sel_folder_button.clicked.connect(
            lambda: self.set_text_button_click(self.sel_folder_text, "Observe_Folder"))

        sel_folder_layout.addWidget(sel_folder_label)
        sel_folder_layout.addWidget(self.sel_folder_text)
        sel_folder_layout.addWidget(sel_folder_button)

        return sel_folder_layout

    def pross_folder_layout(self) -> QHBoxLayout:
        '''Define the layout for processing the selected folder.

        Returns
        -------
        pross_folder_layout : QHBoxLayout
            Layout for processing the selected folder.
        '''

        pross_folder_layout = QHBoxLayout()
        pross_folder_label = QLabel("Folder for processed files:")
        pross_folder_label.setFixedWidth(170)
        self.pross_folder_text = QLineEdit()
        self.pross_folder_text.setFixedWidth(350)
        self.pross_folder_text.setReadOnly(True)
        pross_folder_button = QPushButton("Select Folder")
        pross_folder_button.setFixedWidth(100)

        pross_folder_button.clicked.connect(
            lambda: self.set_text_button_click(self.pross_folder_text, "Processed_Folder"))

        pross_folder_layout.addWidget(pross_folder_label)
        pross_folder_layout.addWidget(self.pross_folder_text)
        pross_folder_layout.addWidget(pross_folder_button)

        return pross_folder_layout

    def not_app_folder_layout(self) -> QHBoxLayout:
        '''Define the layout for the folder where the not applicable files will be saved.

        Returns
        -------
        not_app_folder_layout : QHBoxLayout
            Layout for the folder where the not applicable files will be saved.
        '''

        not_app_folder_layout = QHBoxLayout()
        not_app_folder_label = QLabel("Folder for not applicable files:")
        not_app_folder_label.setFixedWidth(170)
        self.not_app_folder_text = QLineEdit()
        self.not_app_folder_text.setFixedWidth(350)
        self.not_app_folder_text.setReadOnly(True)
        not_app_folder_button = QPushButton("Select Folder")
        not_app_folder_button.setFixedWidth(100)

        not_app_folder_button.clicked.connect(
            lambda: self.set_text_button_click(self.not_app_folder_text, "Not_Applicable_Folder"))

        not_app_folder_layout.addWidget(not_app_folder_label)
        not_app_folder_layout.addWidget(self.not_app_folder_text)
        not_app_folder_layout.addWidget(not_app_folder_button)

        return not_app_folder_layout

    def file_layout(self) -> QHBoxLayout:
        '''Define the layout for the file type to be observed.

        Returns
        -------
        file_layout : QHBoxLayout
            Layout for the file type to be observed.
        '''

        file_layout = QHBoxLayout()
        file_label = QLabel("File to be observed:")
        file_label.setFixedWidth(170)
        self.file_text = QLineEdit()
        self.file_text.setFixedWidth(350)
        self.file_text.setReadOnly(True)
        file_button = QPushButton("Select File")
        file_button.setFixedWidth(100)

        file_button.clicked.connect(
            lambda: self.select_file_button_click(self.file_text, "Consolidation_File"))

        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_text)
        file_layout.addWidget(file_button)

        return file_layout

    def exec_layout(self) -> QVBoxLayout:
        '''Define the layout for the duplication option.

        Returns
        -------
        exec_layout : QHBoxLayout
            Layout for the duplication option.
        '''

        exec_layout = QVBoxLayout()

        check_box_layout = QHBoxLayout()
        check_box_label = QLabel("Allow duplicate files:")
        check_box_label.setFixedWidth(395)
        check_box_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.check_box = QCheckBox()
        self.check_box.setFixedWidth(305)
        check_box_layout.addWidget(check_box_label)
        check_box_layout.addWidget(self.check_box)

        button_layout = QHBoxLayout()
        self.exec_button = QPushButton("Execute")
        self.exec_button.setFixedWidth(100)
        self.exec_button.clicked.connect(self.execute_button_click)
        button_layout.addWidget(self.exec_button)

        state_layout = QHBoxLayout()
        self.state_label = QLabel("State: Idle")
        self.state_label.setFixedWidth(150)
        self.state_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(0)
        self.progress_bar.setFixedWidth(400)
        state_layout.addWidget(self.state_label)
        state_layout.addWidget(self.progress_bar)

        exec_layout.addLayout(check_box_layout)
        exec_layout.addLayout(button_layout)
        exec_layout.addLayout(state_layout)

        return exec_layout

    def set_text_button_click(self, text_line: QLineEdit, field_name: str):
        '''Open a dialog to select a folder and set the text_line its dir.

        Parameters
        ----------
        text_line : QLineEdit
            The QLineEdit widget to be used in the function.
        '''

        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if not folder:
            return
        text_line.setText(folder)
        self.edit_json_field(field_name, folder)

    def select_file_button_click(self, text_line: QLineEdit, field_name: str):
        '''Open a dialog to select a file and set the text_line its dir.'''

        file = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xls? *.xls)")
        text_line.setText(file[0])
        self.edit_json_field(field_name, file[0])

    def execute_button_click(self):
        '''Execute the consolidation process.'''

        if self.active:
            self.end_process()
            return

        self.initiate_process()

    def initiate_process(self):
        '''Initiate the consolidation process.'''

        self.exec_button.setText("Stop")
        self.state_label.setText("State: Processing")
        self.active = True
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)

        self.watcher = Watcher(self)
        self.watcher.comm_signal.connect(self.show_progress)
        if self.watcher.validate_inputs():
            self.watcher.start()
        else:
            self.show_error_message("Invalid Inputs", self.watcher.last_err)

    def end_process(self):
        '''End the consolidation process.'''

        self.progress_bar.setMaximum(0)
        self.progress_bar.setValue(0)
        self.exec_button.setText("Execute")
        self.state_label.setText("State: Idle")
        self.active = False

        self.watcher.quit()
        self.watcher.wait()

    def show_progress(self, message: str):
        '''Show the progress of the consolidation process.'''

        self.progress_bar.setMaximum(0)
        self.state_label.setText(message)

    def show_error_message(self, title: str, message: str):
        '''Show an error message box.

        Parameters
        ----------
        title : str
            The title of the message box.
        '''

        message_box = QMessageBox()
        message_box.setWindowTitle(title)
        message_box.setText(message)
        message_box.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
