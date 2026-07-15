import os
import pandas as pd
import openpyxl
from PySide6.QtCore import QThread, Signal
from template_mapper import TemplateMapper


def process_single_row(args):
    """
    Worker function to process a single antiquity row.
    Runs in a child process.
    args: tuple of (row_dict, template_path, output_folder, index)
    """
    row_dict, template_path, output_folder, index = args
    try:
        # Load the template workbook
        wb = openpyxl.load_workbook(template_path)
        
        # Instantiate mapper and fill record
        mapper = TemplateMapper(wb)
        record = pd.Series(row_dict)
        mapper.fill_record(record)
        
        # Get identifier values for naming
        accession_no = str(row_dict.get("Accession/Registration No", ""))
        if not accession_no or accession_no.lower() == "nan" or not accession_no.strip():
            accession_no = str(row_dict.get("Field Accession No.", ""))
        if not accession_no or accession_no.lower() == "nan" or not accession_no.strip():
            accession_no = str(row_dict.get("NMMA No.", ""))
        title = str(row_dict.get("Title of the object", ""))
        
        # Helper to sanitize string for safe filenames
        def sanitize(val):
            val = str(val).strip()
            if not val or val.lower() == "nan":
                return ""
            # Replace invalid filename characters
            for char in r'\/:*?"<>|':
                val = val.replace(char, "_")
            return val

        san_acc = sanitize(accession_no)
        san_title = sanitize(title)
        
        # Build filename: {index:05d}_{AccessionNo}_{Title}.xlsx
        parts = [f"{index:05d}"]
        if san_acc:
            parts.append(san_acc)
        if san_title:
            # truncate title to 30 chars to avoid hitting path length limits
            parts.append(san_title[:30])
            
        filename = "_".join(parts) + ".xlsx"
        output_file = os.path.join(output_folder, filename)
        
        # Save filled workbook
        wb.save(output_file)
        wb.close()
        
        return {
            "index": index,
            "success": True,
            "output_file": output_file,
            "error": None
        }
    except Exception as e:
        import traceback
        return {
            "index": index,
            "success": False,
            "output_file": None,
            "error": f"{str(e)}\n{traceback.format_exc()}"
        }


def process_chunk(args):
    """
    Worker function to process a chunk of antiquity rows into a single multi-sheet workbook.
    Runs in a child process.
    args: tuple of (rows_list, template_path, output_folder, chunk_index, start_idx, end_idx)
    """
    rows_list, template_path, output_folder, chunk_index, start_idx, end_idx = args
    try:
        # Load the template workbook
        wb = openpyxl.load_workbook(template_path)
        
        # Keep the first sheet as the active template to clone from
        template_sheet = wb.active
        
        for i, row_dict in enumerate(rows_list):
            row_idx = start_idx + i
            
            # Copy template sheet
            ws_copy = wb.copy_worksheet(template_sheet)
            
            # Resolve name for the sheet: index + accession/title fallback
            accession_no = str(row_dict.get("Accession/Registration No", ""))
            if not accession_no or accession_no.lower() == "nan" or not accession_no.strip():
                accession_no = str(row_dict.get("Field Accession No.", ""))
            if not accession_no or accession_no.lower() == "nan" or not accession_no.strip():
                accession_no = str(row_dict.get("NMMA No.", ""))
            title = str(row_dict.get("Title of the object", ""))
            
            # Sanitize for sheet name (max 31 chars, no invalid chars: \ / ? * : [ ])
            def sanitize_sheet_name(val):
                val = str(val).strip()
                if not val or val.lower() == "nan":
                    return ""
                for char in r'\/?*::[]':
                    val = val.replace(char, "_")
                return val

            san_acc = sanitize_sheet_name(accession_no)
            san_title = sanitize_sheet_name(title)
            
            # Form sheet name: e.g. "00001_Toy object"
            sheet_name_parts = [f"{row_idx:05d}"]
            if san_acc:
                sheet_name_parts.append(san_acc)
            elif san_title:
                sheet_name_parts.append(san_title)
                
            sheet_name = "_".join(sheet_name_parts)[:31] # strict 31-char limit in Excel
            
            # Ensure sheet name is unique in this workbook (in case of duplicate indices/names)
            base_sheet_name = sheet_name
            counter = 1
            while sheet_name in wb.sheetnames:
                suffix = f"_{counter}"
                sheet_name = base_sheet_name[:31 - len(suffix)] + suffix
                counter += 1
                
            ws_copy.title = sheet_name
            
            # Fill the data in the copied sheet
            mapper = TemplateMapper(wb)
            mapper.sheet = ws_copy
            
            record = pd.Series(row_dict)
            mapper.fill_record(record)
            
        # Delete the original empty template sheet
        wb.remove(template_sheet)
        
        # Build workbook filename: NMMA_Forms_{start_idx:05d}_to_{end_idx:05d}.xlsx
        filename = f"NMMA_Forms_{start_idx:05d}_to_{end_idx:05d}.xlsx"
        output_file = os.path.join(output_folder, filename)
        
        # Save workbook
        wb.save(output_file)
        wb.close()
        
        return {
            "chunk_index": chunk_index,
            "success": True,
            "output_file": output_file,
            "error": None
        }
    except Exception as e:
        import traceback
        return {
            "chunk_index": chunk_index,
            "success": False,
            "output_file": None,
            "error": f"{str(e)}\n{traceback.format_exc()}"
        }


class ExcelEngine:
    """
    Retained for backward compatibility.
    Runs the legacy sequential generator for the first row.
    """
    def __init__(self, book_path, template_path, output_folder):
        self.book_path = book_path
        self.template_path = template_path
        self.output_folder = output_folder

    def generate_first_form(self):
        df = pd.read_excel(self.book_path)
        record = df.iloc[0]
        wb = openpyxl.load_workbook(self.template_path)
        mapper = TemplateMapper(wb)
        mapper.fill_record(record)
        os.makedirs(self.output_folder, exist_ok=True)
        output_file = os.path.join(self.output_folder, "NMMA_0001.xlsx")
        wb.save(output_file)
        wb.close()
        return output_file


class GeneratorWorker(QThread):
    """
    Worker thread that handles high-performance parallel generation of Excel forms
    (either separate single-form files or grouped multi-sheet workbooks)
    and sequential export of PDF forms in the background.
    """
    progress_updated = Signal(int, int)  # completed, total
    status_updated = Signal(str)
    finished = Signal(int, int, str)     # success_count, error_count, summary_msg
    error_occurred = Signal(str)

    def __init__(self, book_path, template_path, output_folder, output_format, group_size=1):
        super().__init__()
        self.book_path = book_path
        self.template_path = template_path
        self.output_folder = output_folder
        self.output_format = output_format
        self.group_size = group_size
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            # 1. Read Master Data
            self.status_updated.emit("Reading master data file...")
            try:
                df = pd.read_excel(self.book_path)
            except Exception as e:
                self.error_occurred.emit(f"Failed to read master data Excel file:\n{e}")
                return

            total_rows = len(df)
            if total_rows == 0:
                self.error_occurred.emit("The Master data Excel sheet is empty.")
                return

            os.makedirs(self.output_folder, exist_ok=True)

            import multiprocessing
            cpu_count = multiprocessing.cpu_count()
            num_workers = max(1, cpu_count - 1)
            
            success_count = 0
            error_count = 0
            generated_excel_files = []

            # ---------------------------------------------
            # Grouped Workbooks Mode (group_size > 1)
            # ---------------------------------------------
            if self.group_size > 1:
                # Build grouped tasks list
                chunks = []
                chunk_index = 1
                for start_idx in range(0, total_rows, self.group_size):
                    end_idx = min(start_idx + self.group_size, total_rows)
                    chunk_df = df.iloc[start_idx:end_idx]
                    
                    rows_list = [row.to_dict() for _, row in chunk_df.iterrows()]
                    chunks.append((
                        rows_list,
                        self.template_path,
                        self.output_folder,
                        chunk_index,
                        start_idx + 1,
                        end_idx
                    ))
                    chunk_index += 1

                total_chunks = len(chunks)
                self.status_updated.emit(f"Spawning parallel processes for {total_chunks} grouped workbooks...")
                
                pool = multiprocessing.Pool(processes=num_workers)
                results = pool.imap_unordered(process_chunk, chunks)
                
                for res in results:
                    if self._is_cancelled:
                        pool.terminate()
                        pool.join()
                        self.status_updated.emit("Generation cancelled by user.")
                        self.finished.emit(success_count, error_count, "Generation cancelled by user.")
                        return

                    if res["success"]:
                        success_count += 1
                        generated_excel_files.append(res["output_file"])
                    else:
                        error_count += 1
                        print(f"Error processing chunk {res['chunk_index']}: {res['error']}")
                        
                    completed_chunks = success_count + error_count
                    self.progress_updated.emit(completed_chunks, total_chunks)
                    self.status_updated.emit(f"Generated {completed_chunks}/{total_chunks} grouped workbooks...")

                pool.close()
                pool.join()

            # ---------------------------------------------
            # Individual Files Mode (group_size <= 1)
            # ---------------------------------------------
            else:
                # Build tasks list
                tasks = []
                for idx, (_, row) in enumerate(df.iterrows(), start=1):
                    row_dict = row.to_dict()
                    tasks.append((row_dict, self.template_path, self.output_folder, idx))

                self.status_updated.emit(f"Spawning parallel processes for {total_rows} antiquity forms...")
                
                pool = multiprocessing.Pool(processes=num_workers)
                results = pool.imap_unordered(process_single_row, tasks)
                
                for res in results:
                    if self._is_cancelled:
                        pool.terminate()
                        pool.join()
                        self.status_updated.emit("Generation cancelled by user.")
                        self.finished.emit(success_count, error_count, "Generation cancelled by user.")
                        return

                    if res["success"]:
                        success_count += 1
                        generated_excel_files.append(res["output_file"])
                    else:
                        error_count += 1
                        print(f"Error processing row index {res['index']}: {res['error']}")

                    processed = success_count + error_count
                    self.progress_updated.emit(processed, total_rows)
                    self.status_updated.emit(f"Generated {processed}/{total_rows} Excel forms...")

                pool.close()
                pool.join()

            # 4. PDF Export phase if selected
            if self.output_format in ["PDF", "Excel + PDF"] and not self._is_cancelled:
                self.status_updated.emit("Initializing Microsoft Excel for PDF conversion...")
                self.progress_updated.emit(0, len(generated_excel_files))

                from pdf_exporter import ExcelPDFConverter
                converter = ExcelPDFConverter()
                try:
                    converter.start()
                except Exception as e:
                    self.error_occurred.emit(str(e))
                    return

                pdf_success = 0
                pdf_error = 0
                total_pdfs = len(generated_excel_files)

                for idx, xlsx_path in enumerate(generated_excel_files, start=1):
                    if self._is_cancelled:
                        break

                    self.status_updated.emit(f"Converting to PDF ({idx}/{total_pdfs}): {os.path.basename(xlsx_path)}")
                    
                    try:
                        # Construct PDF path
                        pdf_filename = os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf"
                        pdf_path = os.path.join(self.output_folder, pdf_filename)
                        
                        converter.convert(xlsx_path, pdf_path)
                        pdf_success += 1
                        
                        # Clean up Excel file if user ONLY wanted PDF
                        if self.output_format == "PDF":
                            try:
                                os.remove(xlsx_path)
                            except Exception:
                                pass
                    except Exception as e:
                        pdf_error += 1
                        print(f"PDF conversion failed for {xlsx_path}: {e}")

                    self.progress_updated.emit(idx, total_pdfs)

                converter.stop()

                if self._is_cancelled:
                    self.status_updated.emit("Cancelled during PDF export.")
                    self.finished.emit(success_count, error_count, "PDF conversion cancelled by user.")
                    return

                summary = (
                    f"Completed successfully!\n\n"
                    f"Excel Files Generated: {success_count} (Failed: {error_count})\n"
                    f"PDF Files Generated: {pdf_success} (Failed: {pdf_error})"
                )
                self.finished.emit(success_count, error_count, summary)

            else:
                summary = (
                    f"Completed successfully!\n\n"
                    f"Excel Files Generated: {success_count} (Failed: {error_count})"
                )
                self.finished.emit(success_count, error_count, summary)

        except Exception as e:
            import traceback
            self.error_occurred.emit(f"Unexpected worker error:\n{e}\n{traceback.format_exc()}")