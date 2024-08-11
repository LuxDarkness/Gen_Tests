'''This module is used to handle the file events'''

import os
from watchdog.events import FileSystemEventHandler

class FileHandler(FileSystemEventHandler):
    '''This class is used to handle the file events'''

    def __init__(self, watcher):
        self.watcher = watcher

    def on_created(self, event: FileSystemEventHandler):
        '''This method is used to handle the file creation event'''

        if event.is_directory:
            return

        if os.path.basename(event.src_path).startswith("~$"):
            return

        if not event.src_path.endswith((".xlsx", ".xlsb", ".xlsm", ".xls")):
            self.watcher.comm_signal.emit(
                f"State: Moving new file '{os.path.basename(event.src_path)}'")
            self.watcher.move_not_applicable(event.src_path)
            return

        self.watcher.comm_signal.emit(
            f"State: Processing new file '{os.path.basename(event.src_path)}'")
        self.watcher.process_file(event.src_path)
