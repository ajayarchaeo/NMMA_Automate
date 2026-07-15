class TemplateMapper:

    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = workbook.active

    def fill_record(self, record):
        """
        Fills the NMMA template with data from one row of Book1.xlsx
        """

        # Safe function to avoid KeyErrors and NaN values
        def value(column_name):
            if column_name in record.index:
                v = record[column_name]
                if str(v) == "nan" or v is None:
                    return ""
                return str(v).strip()
            # Case insensitive check as fallback
            for col in record.index:
                if str(col).lower() == str(column_name).lower():
                    v = record[col]
                    if str(v) == "nan" or v is None:
                        return ""
                    return str(v).strip()
            return ""

        # Helper to only write to cell if we have a non-empty value
        # (this preserves defaults pre-filled in the template)
        def fill_cell(cell_ref, column_name):
            val = value(column_name)
            if val != "":
                self.sheet[cell_ref] = val

        # -----------------------------
        # Fill NMMA Template Cell Mappings
        # -----------------------------

        # Title and Type
        fill_cell("C3", "Title of the object")
        fill_cell("C4", "Type of the object")
        
        # Period, Dynasty, and Provenance
        fill_cell("C5", "Date/Period")
        fill_cell("C6", "Dynasty/Style")
        fill_cell("C7", "Provenance")
        
        # Material
        fill_cell("C8", "Material")

        # Measurement / Weight (with smart fallback combinations)
        meas_val = value("Measurement/weight")
        if not meas_val:
            meas_parts = []
            dia = value("Diameter")
            if dia:
                meas_parts.append(f"Dia: {dia} cm")
            thk = value("Thickness")
            if thk:
                meas_parts.append(f"Thk: {thk} cm")
            wt = value("Weight (g)")
            if wt:
                meas_parts.append(f"Wt: {wt} g")
            if meas_parts:
                meas_val = ", ".join(meas_parts)
                
        if meas_val:
            self.sheet["C9"] = meas_val

        # Description and Marks
        fill_cell("C10", "Description")
        fill_cell("C11", "Identification Marks")
        
        # Condition, Photo, Location, and State
        fill_cell("C12", "Condition")
        fill_cell("C13", "Photograph")
        fill_cell("C14", "Location at the Museum")
        fill_cell("C15", "State/UT")
        
        # Accession/Registration No (Cell C16)
        acc_val = value("Accession/Registration No")
        if not acc_val:
            acc_val = value("NMMA No.")
        if acc_val:
            self.sheet["C16"] = acc_val
        fill_cell("C17", "Source of Acquisition")
        fill_cell("C18", "National documentation Number")
        fill_cell("C19", "Published References")

        # Remarks (with smart Trench/Depth fallback)
        remarks_val = value("Remarks")
        trench = value("Trench")
        depth = value("Depth (cm)")
        
        context_parts = []
        if trench:
            context_parts.append(f"Trench: {trench}")
        if depth:
            context_parts.append(f"Depth: {depth}")
            
        if context_parts:
            context_prefix = ", ".join(context_parts)
            if remarks_val:
                remarks_val = f"{context_prefix}. {remarks_val}"
            else:
                remarks_val = context_prefix
                
        if remarks_val:
            self.sheet["C20"] = remarks_val

        # Recording details
        fill_cell("C21", "Date of Recording")
        fill_cell("C22", "Recorded by")

        return self.workbook