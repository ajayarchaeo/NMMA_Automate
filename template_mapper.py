import os
from openpyxl.utils import coordinate_to_tuple, get_column_letter


def offset_cell(cell_ref, row_offset=0, col_offset=0):
    """
    Returns a cell coordinate string shifted by row_offset and col_offset.
    For example: offset_cell("C3", row_offset=24, col_offset=5) returns "H27".
    """
    row, col = coordinate_to_tuple(cell_ref)
    new_row = row + row_offset
    new_col = col + col_offset
    return f"{get_column_letter(new_col)}{new_row}"


class TemplateMapper:

    def __init__(self, workbook):
        self.workbook = workbook
        self.sheet = workbook.active

    def fill_record(self, record, mapping_dict, row_offset=0, col_offset=0, photo_path=None, photo_width=220, photo_height=200):
        """
        Fills the NMMA template dynamically using mapping_dict.
        Supports shifting target cell coordinates dynamically.
        Inserts and scales a photograph at cell C13 with user-defined photo_width and photo_height constraints.
        mapping_dict is a dict: {cell_ref: column_name}
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

        # Shift cell coordinate dynamically
        def fill_cell(cell_ref, column_name):
            val = value(column_name)
            if val != "":
                target_cell = offset_cell(cell_ref, row_offset, col_offset)
                self.sheet[target_cell] = val

        # -----------------------------
        # Fill mapped cells
        # -----------------------------
        for cell_ref, col_name in mapping_dict.items():
            if not col_name or col_name == "(Use Template Default)":
                continue
            fill_cell(cell_ref, col_name)

        # -----------------------------
        # Photograph Insertion (Cell C13)
        # -----------------------------
        if photo_path and os.path.exists(photo_path):
            from openpyxl.drawing.image import Image as OpenpyxlImage
            try:
                img = OpenpyxlImage(photo_path)
                
                # Dynamic constraints from the user settings
                max_w = photo_width
                max_h = photo_height
                orig_w = img.width
                orig_h = img.height
                
                # Scale while preserving original aspect ratio
                scale = min(max_w / orig_w, max_h / orig_h)
                if scale < 1.0:
                    img.width = int(orig_w * scale)
                    img.height = int(orig_h * scale)
                    
                target_cell = offset_cell("C13", row_offset, col_offset)
                self.sheet.add_image(img, target_cell)
            except Exception as e:
                print(f"Error inserting image {photo_path} at {target_cell}: {e}")

        # -----------------------------
        # Smart Fallbacks (Applied after direct mappings)
        # -----------------------------
        
        # 1. Accession/Registration No fallback (Cell C16)
        c16_target = offset_cell("C16", row_offset, col_offset)
        acc_col = mapping_dict.get("C16")
        acc_val = ""
        if acc_col and acc_col != "(Use Template Default)":
            acc_val = value(acc_col)
            
        # If the master column value was empty or not set, run fallback
        if not acc_val:
            acc_val = value("Field Accession No.")
            if not acc_val:
                acc_val = value("NMMA No.")
            if acc_val:
                self.sheet[c16_target] = acc_val

        # 2. Measurement/Weight fallback (Cell C9)
        c9_target = offset_cell("C9", row_offset, col_offset)
        meas_col = mapping_dict.get("C9")
        meas_val = ""
        if meas_col and meas_col != "(Use Template Default)":
            meas_val = value(meas_col)
            
        if not meas_val:
            # Try to build from Length, Diameter, Thickness, Weight (g)
            meas_parts = []
            length = value("Length")
            if length:
                meas_parts.append(f"L: {length} cm")
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
                self.sheet[c9_target] = ", ".join(meas_parts)

        # 3. Remarks fallback (Cell C20)
        c20_target = offset_cell("C20", row_offset, col_offset)
        remarks_col = mapping_dict.get("C20")
        remarks_val = ""
        if remarks_col and remarks_col != "(Use Template Default)":
            remarks_val = value(remarks_col)
            
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
            self.sheet[c20_target] = remarks_val

        return self.workbook