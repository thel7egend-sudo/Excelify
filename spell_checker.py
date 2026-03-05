import re

from spylls.hunspell import Dictionary


class SpellChecker:
    def __init__(self, dictionary_path="dictionaries/en_US"):
        self._dictionary = None
        self._word_pattern = re.compile(r"[A-Za-z']+")
        try:
            self._dictionary = Dictionary.from_files(dictionary_path)
        except Exception:
            self._dictionary = None

    def is_correct(self, word):
        if not word or self._dictionary is None:
            return True
        return bool(self._dictionary.lookup(word.lower()))

    def misspelled_ranges(self, text):
        if not text or self._dictionary is None:
            return []

        ranges = []
        for match in self._word_pattern.finditer(text):
            word = match.group(0)
            if not self.is_correct(word):
                ranges.append((match.start(), match.end(), word))
        return ranges
