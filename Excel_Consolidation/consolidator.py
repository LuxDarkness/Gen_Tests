'''This module is used to consolidate the files'''

import os
import shutil
import datetime
import xlwings as xw

class Consolidator:
    '''This class is used to consolidate the files'''

    def __init__(self, base_file: str, report_sheet: str, duplicates: bool):
        self.consolidator_file = base_file
        self.report_sheet = report_sheet
        self.enable_duplicates = duplicates
        self.moved = False
        self.is_duplicate = False

    def consolidate(self, file_path: str):
        '''This method is used to consolidate the files

        Parameters
        ----------
        file_path: str
            The file path of the file to be consolidated
        '''

        app = None
        try:
            app = xw.App(visible=False)
            target_wb = xw.Book(self.consolidator_file)
            source_wb = xw.Book(file_path)

            if not self.report_sheet in [sheet.name for sheet in target_wb.sheets]:
                target_wb.sheets.add(self.report_sheet, before=target_wb.sheets[0])

            self.setup_report(target_wb.sheets[self.report_sheet])
            if not self.check_duplicates(target_wb, file_path):
                print(f"File: {file_path} is a duplicate")
                return

            for sheet in source_wb.sheets:
                new_name = self.get_unique_name(target_wb, sheet.name)
                sheet.api.Copy(After=target_wb.sheets[-1].api)
                target_wb.sheets[-1].name = new_name
                self.register_report(target_wb.sheets[self.report_sheet], file_path, new_name)

            source_wb.close()
            target_wb.save()
            target_wb.close()
        except Exception as e:
            print(f"Error: {str(e)}")
        finally:
            if app is not None:
                app.quit()

    def move_excel(self, file_path: str, dest_folder: str):
        '''This method is used to move the excel file to the destination folder'''

        destination = os.path.join(dest_folder, os.path.basename(file_path))
        can_move = True
        if os.path.exists(destination):
            can_move = False
            if os.access(destination, os.W_OK):
                can_move = True
                os.remove(destination)

        if can_move:
            shutil.move(file_path, destination)

        self.moved = can_move

    def get_unique_name(self, target_wb: xw.Book, sheet_name: str) -> str:
        '''This method is used to get a unique name for the sheet'''

        existing_names = [sheet.name for sheet in target_wb.sheets]

        if not sheet_name in existing_names:
            return sheet_name

        suffix = 1
        while f"{sheet_name} ({suffix})" in existing_names:
            suffix += 1

        return f"{sheet_name} ({suffix})"

    def setup_report(self, report_sheet: xw.Sheet):
        '''This method is used to setup the report sheet'''

        report_sheet.range("A1").value = "File Name"
        report_sheet.range("B1").value = "Sheet Name"
        report_sheet.range("C1").value = "Date and Time"
        report_sheet.range("D1").value = "Status"
        report_sheet.range("E1").value = "Is Duplicate File"
        report_sheet.range("F1").value = "Message"

    def register_report(self, report_sheet: xw.Sheet, file_path: str, sheet_name: str,
                        status: str = "Success", msg: str = ""):
        '''This method is used to register each processed file and sheet to the report

        Parameters
        ----------
        report_sheet: xw.Sheet
            The sheet where the report will be registered
        file_path: str
            The file path of the processed file
        sheet_name: str
            The name of the processed sheet
        status: str
            The status of the process
        msg: str
            An additional message of the process
        '''

        row = report_sheet.range('A2')
        if report_sheet.range('A2').value:
            row = report_sheet.range('A' + str(report_sheet.range('A1').end('down').row + 1))

        if not self.moved:
            msg = "Sheets copied but failed to move the file"
        row.value = [os.path.basename(file_path), sheet_name,
                    datetime.datetime.now(), status, self.is_duplicate, msg]

    def register_info(self, file_path: str, status: str, msg: str):
        '''This method is used to register the info of special cases

        Parameters
        ----------
        file_path: str
            The file path of the processed file
        status: str
            The status of the process
        msg: str
            An additional message of the process
        '''

        app = None
        try:
            app = xw.App(visible=False)
            wb = xw.Book(self.consolidator_file)

            if not self.report_sheet in [sheet.name for sheet in wb.sheets]:
                wb.sheets.add(self.report_sheet, before=wb.sheets[0])

            ws = wb.sheets[self.report_sheet]
            self.setup_report(ws)

            row = ws.range('A2')
            if ws.range('A2').value:
                row = ws.range('A' + str(ws.range('A1').end('down').row + 1))

            full_msg = f"File: {os.path.basename(file_path)}\n{msg}"
            row.value = ["-", "N/A", datetime.datetime.now(), status, False, full_msg]

            wb.save()
            wb.close()
        except Exception as e:
            print(f"Error: {str(e)}")
        finally:
            if app is not None:
                app.quit()

    def check_duplicates(self, target_wb: xw.Book, file_path: str) -> bool:
        '''This method is used to check if the file is a duplicate'''

        file_name = os.path.basename(file_path)
        last_row = 2
        if target_wb.sheets[self.report_sheet].range('A2').value:
            last_row = target_wb.sheets[self.report_sheet].range('A1').end('down').row
        files_list = [cell.value for cell in target_wb.sheets[self.report_sheet].range("A2:A"
            + str(last_row)) if cell.value]
        unique_files = set(files_list)

        if self.enable_duplicates:
            if file_name in unique_files:
                self.is_duplicate = True
            return True

        if file_name in unique_files:
            self.is_duplicate = True
            self.register_report(
                target_wb.sheets[self.report_sheet], file_path, "N/A", "Failed", "Duplicate file")
            return False

        return True
