'''This module is used to watch the folder for any new files and consolidate them 
    into a single file'''

import time
import json
import os
import shutil
from PyQt6.QtCore import QThread, pyqtSignal
from watchdog.observers import Observer
from file_handler import FileHandler

class Watcher(QThread):
    '''This class is used to watch the folder for any new files and consolidate 
        them into a single file'''

    comm_signal = pyqtSignal(str)

    def __init__(self, main_window):
        super().__init__()
        self.observer = Observer()

        with open("Resources/config.json", "r", encoding='UTF-8') as file:
            data = json.load(file)

        self.last_err = None
        self.file_path = data["Consolidation_File"]
        self.observe_folder = data["Observe_Folder"]
        self.processed_folder = data["Processed_Folder"]
        self.not_applicable_folder = data["Not_Applicable_Folder"]
        self.main_window = main_window

    def run(self):
        '''This method is used to start the observer and watch the folder for any new files'''

        try:
            self.process_existing_files()
        except PermissionError as e:
            self.last_err = f'While processing existing files, this exception ocurred: {str(e)}'
            self.main_window.show_error("Permission Denied", self.last_err)
            return

        event_handler = FileHandler(self)
        self.observer.schedule(event_handler, self.observe_folder, recursive=False)
        self.observer.start()

        try:
            while self.main_window.active:
                time.sleep(3)
        except Exception:
            pass
        finally:
            self.observer.stop()

        self.observer.join()

    def process_file(self, file_path: str):
        '''This method is used to process the file'''

        print(f"Processing file: {file_path}")

    def move_not_applicable(self, file_path: str):
        '''This method is used to move the file to the not applicable folder'''

        if not self.check_permission(file_path, self.not_applicable_folder):
            print(f"Permission denied for file: {file_path}")
            return

        destination = os.path.join(self.not_applicable_folder, os.path.basename(file_path))
        shutil.move(file_path, destination)

    def validate_inputs(self):
        '''This method is used to validate the inputs'''

        if not os.path.exists(self.observe_folder):
            self.last_err = "Observe folder does not exist"
            return False

        if not os.path.exists(self.processed_folder):
            self.last_err = "Processed folder does not exist"
            return False

        if not os.path.exists(self.not_applicable_folder):
            self.last_err = "Not applicable folder does not exist"
            return False

        if not os.path.exists(self.file_path):
            self.last_err = "Consolidation file does not exist"
            return False

        return True

    def check_permission(self, file_path: str, destination: str):
        '''This method is used to check the permission'''

        return os.access(file_path, os.R_OK) and os.access(destination, os.W_OK)

    def process_existing_files(self):
        '''This method is used to process the existing files'''

        if not os.access(self.observe_folder, os.R_OK):
            print(f"Permission denied for folder: {self.observe_folder}")
            raise PermissionError("Permission denied")

        for file in os.listdir(self.observe_folder):
            file_path = os.path.join(self.observe_folder, file)
            if not os.path.isfile(file_path):
                continue
            if not file_path.endswith((".xlsx", ".xlsb", ".xlsm", ".xls")):
                self.move_not_applicable(file_path)
                continue
            self.process_file(file_path)
