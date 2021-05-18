import os
import openpyxl
import docx


class CUSTOM_TYPES:
    class ALIGNMENT:
        LEFT = 0
        CENTER = 1
        JUSTIFY = 3
        RIGHT = 2


class ParseWord:
    """analyzing the word.doc/docx syntax"""
    def __init__(self, path: str):
        assert len(path) > 0
        self._parsedText: [{"str"}] = []
        self._path = path

    def parse(self):
        doc = docx.Document(self._path)
        for paragraph in doc.paragraphs:

            self._parsedText.append([
                {
                    "text": paragraph.text,
                    "properties": {
                        "alignment": paragraph.alignment,
                        "indents": self.getIndents(paragraph)
                    }

                }
            ])
        return self._parsedText

    def getIndents(self, paragraph) -> {str}:
        """returning indents in pt format"""
        formatting = paragraph.paragraph_format
        before = formatting.space_before
        after = formatting.space_after
        left = formatting.left_indent
        right = formatting.right_indent
        first_line = formatting.first_line_indent

        if before is not None:
            before = before.pt

        if after is not None:
            after = after.pt

        if left is not None:
            left = left.pt

        if right is not None:
            right = right.pt

        if first_line is not None:
            first_line = first_line.pt

        return {
            "before": before,
            "after": after,
            "left": left,
            "right": right,
            "first_line": first_line
        }


class ImportInExcel:
    def __init__(self, text_for_import: [{str}]):
        assert len(text_for_import) > 0
        self.__text_for_import = text_for_import

    def insert(self):
        pass


def main():
    # doc = docx.Document(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'text.docx'))
    pars = ParseWord(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'text.docx'))
    parsedText = pars.parse()
    print(parsedText)
    importer = ImportInExcel(parsedText)
    importer.insert()


if __name__ == '__main__':
    main()

