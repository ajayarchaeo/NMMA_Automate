import os
import threading
import queue
import win32com.client
import pythoncom


class ExcelCOMThread(threading.Thread):
    """
    Dedicated worker thread for handling Microsoft Excel COM operations.
    Keeps COM initialization and execution isolated to prevent cross-apartment threading issues
    and enables timeout checks.
    """

    def __init__(self):
        super().__init__()
        self.task_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.excel = None
        self.daemon = True

    def run(self):
        pythoncom.CoInitialize()
        try:
            # Dispatch Excel
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.result_queue.put(("START_OK", None))
        except Exception as e:
            self.result_queue.put(("START_ERR", str(e)))
            pythoncom.CoUninitialize()
            return

        while True:
            cmd, args = self.task_queue.get()
            if cmd == "CONVERT":
                xlsx_path, pdf_path = args
                wb = None
                try:
                    abs_xlsx = os.path.abspath(xlsx_path)
                    abs_pdf = os.path.abspath(pdf_path)
                    wb = self.excel.Workbooks.Open(abs_xlsx)
                    # 0 is xlTypePDF, export entire workbook to include all sheets
                    wb.ExportAsFixedFormat(0, abs_pdf)
                    self.result_queue.put(("CONVERT_OK", None))
                except Exception as e:
                    self.result_queue.put(("CONVERT_ERR", str(e)))
                finally:
                    if wb:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass
            elif cmd == "STOP":
                try:
                    if self.excel:
                        self.excel.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()
                self.result_queue.put(("STOP_OK", None))
                break


class ExcelPDFConverter:
    """
    Wrapper for Excel PDF conversion. Initiates Excel COM in a separate thread
    and guards against hung states using initialization timeouts.
    """

    def __init__(self):
        self.com_thread = None
        self.initialized = False

    def start(self, timeout=10):
        """
        Starts the COM thread and waits for Microsoft Excel to initialize.
        Raises a detailed RuntimeError if it fails or times out.
        """
        if self.initialized:
            return

        self.com_thread = ExcelCOMThread()
        self.com_thread.start()

        try:
            res, err = self.com_thread.result_queue.get(timeout=timeout)
            if res == "START_OK":
                self.initialized = True
            else:
                raise RuntimeError(
                    f"Failed to start Excel COM application.\n"
                    f"Please verify that Microsoft Excel is installed and configured on this system.\n"
                    f"Details: {err}"
                )
        except queue.Empty:
            raise RuntimeError(
                "Microsoft Excel COM initialization timed out (10s).\n\n"
                "This usually occurs if Microsoft Excel is waiting for a background dialog "
                "(such as a license activation screen, first-run wizard, or update prompt) to be completed.\n\n"
                "Please open Microsoft Excel manually on this machine, close any pop-ups or activation dialogs, "
                "and then run this tool again."
            )

    def convert(self, xlsx_path, pdf_path):
        """
        Requests the COM thread to convert a single Excel file to PDF.
        """
        if not self.initialized or not self.com_thread:
            raise RuntimeError("ExcelPDFConverter is not started. Call start() first.")

        self.com_thread.task_queue.put(("CONVERT", (xlsx_path, pdf_path)))
        res, err = self.com_thread.result_queue.get()
        if res != "CONVERT_OK":
            raise RuntimeError(f"Failed to convert Excel to PDF: {err}")

    def stop(self):
        """
        Gracefully stops the Excel application and shuts down the COM thread.
        """
        if not self.initialized or not self.com_thread:
            return

        try:
            self.com_thread.task_queue.put(("STOP", None))
            self.com_thread.result_queue.get(timeout=5)
        except Exception:
            pass
        finally:
            self.com_thread = None
            self.initialized = False
