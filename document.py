class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}   # {(row, col): value}
        self.row_heights = {}  # {row: height}
        self.col_widths = {}   # {col: width}

    def to_dict(self):
        return {
            "name": self.name,
            "cells": {
                f"{r},{c}": v for (r, c), v in self.cells.items()
            },
            "row_heights": {str(r): h for r, h in self.row_heights.items()},
            "col_widths": {str(c): w for c, w in self.col_widths.items()},
        }

    @staticmethod
    def from_dict(data):
        sheet = Sheet(data["name"])
        for key, value in data.get("cells", {}).items():
            r, c = map(int, key.split(","))
            sheet.cells[(r, c)] = value
        sheet.row_heights = {
            int(r): h for r, h in data.get("row_heights", {}).items()
        }
        sheet.col_widths = {
            int(c): w for c, w in data.get("col_widths", {}).items()
        }
        return sheet


class Document:
    def __init__(self, name):
        self.name = name
        self.sheets = [Sheet("Sheet1")]
        self.active_sheet_index = 0

    @property
    def active_sheet(self):
        return self.sheets[self.active_sheet_index]

    def to_dict(self):
        return {
            "name": self.name,
            "active_sheet_index": self.active_sheet_index,
            "sheets": [sheet.to_dict() for sheet in self.sheets],
        }

    @staticmethod
    def from_dict(data):
        doc = Document(data["name"])
        doc.sheets = [
            Sheet.from_dict(s) for s in data.get("sheets", [])
        ]
        doc.active_sheet_index = data.get("active_sheet_index", 0)

        # safety: ensure at least one sheet
        if not doc.sheets:
            doc.sheets = [Sheet("Sheet1")]
            doc.active_sheet_index = 0

        return doc
