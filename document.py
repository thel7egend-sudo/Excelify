class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}   # {(row, col): value}

    def to_dict(self):
        return {
            "name": self.name,
            "cells": {
                f"{r},{c}": v for (r, c), v in self.cells.items()
            }
        }

    @staticmethod
    def from_dict(data):
        sheet = Sheet(data["name"])
        for key, value in data.get("cells", {}).items():
            r, c = map(int, key.split(","))
            sheet.cells[(r, c)] = value
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
