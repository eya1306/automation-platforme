"""Call-tree comparison, lifted from `comp_call_tree/ml_comparator.py`.

The `SimpleCallTreeComparator` class below is the upstream one verbatim; only
the tkinter front-end that shared the file has been left behind. It already
returns a summary dict from `create_comparison_report`, which is what the run
panel reports.

The `print` calls inside it are harmless here — the job runner captures stdout
per run, so they end up in the run log rather than on a terminal.
"""

from __future__ import annotations

import difflib
import pathlib

import openpyxl
from openpyxl.styles import Font, PatternFill


class SimpleCallTreeComparator:
    def __init__(self):
        """Initialise le comparateur simple"""
        # Seuil de similarité : 0.0 à 1.0
        # 0.85 = 85% de ressemblance pour détecter un renommage
        self.similarity_threshold = 0.85

    def extract_rows(self, ws):
        """Extrait chaque ligne sous forme de texte concaténé"""
        rows = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            values = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
            if values:
                rows.append((row_idx, " | ".join(values)))
        return rows

    def _similarity(self, a, b):
        """Calcule la similarité entre deux chaînes (0.0 → 1.0)"""
        return difflib.SequenceMatcher(None, a, b).ratio()

    def find_similar_changes(self, before_rows, after_rows):
        """
        Détecte les modifications entre lignes :
        - inchangées : texte identique
        - supprimées : présentes avant, pas après
        - ajoutées   : présentes après, pas avant
        - modifiées  : lignes similaires (similarité > seuil)
        """
        remaining_after = list(after_rows)
        unchanged = []
        deleted = []
        modified = []

        # 1. Correspondances exactes
        for before_row in before_rows:
            found = False
            for idx, after_row in enumerate(remaining_after):
                if before_row[1] == after_row[1]:
                    unchanged.append((before_row, after_row))
                    del remaining_after[idx]
                    found = True
                    break
            if not found:
                deleted.append(before_row)

        added = list(remaining_after)

        # 2. Parmi supprimées + ajoutées → détecter les renommages/modifications
        if deleted and added:
            print("Recherche de renommages par similarité...")
            matched_deleted = set()
            matched_added = set()

            for i, del_row in enumerate(deleted):
                best_j = None
                best_score = 0.0
                for j, add_row in enumerate(added):
                    if j in matched_added:
                        continue
                    score = self._similarity(del_row[1], add_row[1])
                    if score > best_score:
                        best_score = score
                        best_j = j
                if best_j is not None and best_score >= self.similarity_threshold:
                    modified.append({
                        'before': deleted[i],
                        'after': added[best_j],
                        'similarity': best_score
                    })
                    matched_deleted.add(i)
                    matched_added.add(best_j)

            deleted = [row for idx, row in enumerate(deleted) if idx not in matched_deleted]
            added   = [row for idx, row in enumerate(added)   if idx not in matched_added]

        return {
            'deleted':   deleted,
            'added':     added,
            'modified':  modified,
            'unchanged': unchanged
        }
    
    def _load_sheet(self, file_path, sheet_name):
        wb = openpyxl.load_workbook(file_path)
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        return wb.active
    
    def _copy_sheet_with_highlights(self, src_ws, dst_ws, highlight_rows, fill_color="F44336", font_color="FFFFFF"):
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        font = Font(color=font_color)
        
        max_row = src_ws.max_row or 1
        max_col = src_ws.max_column or 1
        
        for row_num in range(1, max_row + 1):
            highlight = row_num in highlight_rows
            for col_num in range(1, max_col + 1):
                src_cell = src_ws.cell(row=row_num, column=col_num)
                dst_cell = dst_ws.cell(row=row_num, column=col_num, value=src_cell.value)
                if highlight:
                    dst_cell.fill = fill
                    dst_cell.font = font
        
        for col_letter, dim in src_ws.column_dimensions.items():
            if dim.width:
                dst_ws.column_dimensions[col_letter].width = dim.width
    
    def _get_changed_row_indices(self, changes):
        before_changed_rows = set()
        after_changed_rows = set()
        
        for row_index, _ in changes['deleted']:
            before_changed_rows.add(row_index)
        
        for row_index, _ in changes['added']:
            after_changed_rows.add(row_index)
        
        for mod in changes['modified']:
            before_changed_rows.add(mod['before'][0])
            after_changed_rows.add(mod['after'][0])
        
        return before_changed_rows, after_changed_rows
    
    def create_comparison_report(self, before_file, after_file=None, output_file=None,
                                 before_sheet="Call Tree", after_sheet="Call Tree (Apres modif)"):
        """Crée le rapport: 2 sheets (before/after) avec toutes les lignes, changements en rouge"""
        
        if after_file is None:
            print(f"Lecture: {before_file}")
            wb = openpyxl.load_workbook(before_file)
            before_ws = wb[before_sheet]
            after_ws = wb[after_sheet]
        else:
            print(f"Lecture BEFORE: {before_file}")
            print(f"Lecture AFTER: {after_file}")
            before_ws = self._load_sheet(before_file, before_sheet)
            after_ws = self._load_sheet(after_file, after_sheet)
        
        before_rows = self.extract_rows(before_ws)
        after_rows = self.extract_rows(after_ws)
        
        print(f"Before: {len(before_rows)} lignes")
        print(f"After: {len(after_rows)} lignes")
        
        changes = self.find_similar_changes(before_rows, after_rows)
        before_changed_rows, after_changed_rows = self._get_changed_row_indices(changes)
        
        report_wb = openpyxl.Workbook()
        default_sheet = report_wb.active
        report_wb.remove(default_sheet)
        
        ws_before = report_wb.create_sheet("before")
        ws_after = report_wb.create_sheet("after")
        
        self._copy_sheet_with_highlights(before_ws, ws_before, before_changed_rows, fill_color="F44336", font_color="FFFFFF")
        self._copy_sheet_with_highlights(after_ws, ws_after, after_changed_rows, fill_color="4CAF50", font_color="FFFFFF")
        
        if output_file is None:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = pathlib.Path(before_file).stem
            output_file = pathlib.Path(before_file).with_name(f"{base_name}_ML_{timestamp}.xlsx")
        else:
            output_file = pathlib.Path(output_file)
        
        output_file = pathlib.Path(output_file)
        
        try:
            report_wb.save(output_file)
        except PermissionError:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            alternate = output_file.with_name(f"{output_file.stem}_{timestamp}{output_file.suffix}")
            print(f"Permission denied writing {output_file}, saving to {alternate}")
            report_wb.save(alternate)
            output_file = alternate
        
        print(f"Rapport généré: {output_file}")
        
        return {
            'deleted': len(changes['deleted']),
            'added': len(changes['added']),
            'modified': len(changes['modified']),
            'unchanged': len(changes['unchanged']),
            'output_file': str(output_file)
        }
