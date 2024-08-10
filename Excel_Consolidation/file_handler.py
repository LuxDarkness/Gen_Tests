'''This module is used to handle the file events'''

from watchdog.events import FileSystemEventHandler

class FileHandler(FileSystemEventHandler):
    '''This class is used to handle the file events'''

    def __init__(self, watcher):
        self.watcher = watcher

    def on_created(self, event: FileSystemEventHandler):
        '''This method is used to handle the file creation event'''

        if event.is_directory:
            return

        self.watcher.comm_signal.emit(f"State: Processing new file '{event.src_path}'")

        if not event.src_path.endswith((".xlsx", ".xlsb", ".xlsm", ".xls")):
            self.watcher.move_not_applicable(event.src_path)
            return

        self.watcher.process_file(event.src_path)
