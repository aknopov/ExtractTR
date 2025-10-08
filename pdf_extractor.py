import fitz # PyMuPDF

class PdfExtractor:
    def __init__(self, pdf_file: str):
        self.doc = fitz.open(pdf_file)

    def __del__(self):
        if (self.doc != None):
            self.doc.close()

    def num_pages(self) -> int:
        return self.doc.page_count

    def read_page(self, page_num: int) -> str:
        page = self.doc[page_num]
        return page.get_text(sort=True)

    def find_values(self, page_num: int, key: str, num_vals = 1) -> list:
        """Finds words on a page that are preceded by a parsed 'key' words."""
        key_words = key.split()
        page = self.doc[page_num]
        page_boxes = page.get_text("words", sort=True)
        # As per https://pymupdf.readthedocs.io/en/latest/textpage.html#TextPage.extractWORDS
        page_words = [row[4] for row in page_boxes]

        num_keys = len(key_words)
        num_words = len(page_words)

        for i in range(num_words - num_keys + 1):
            if page_words[i : i + num_keys] == key_words:
                return page_words[i + num_keys : i + num_keys + num_vals]

        return None
