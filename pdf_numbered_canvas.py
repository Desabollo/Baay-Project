from reportlab.pdfgen import canvas
from reportlab.lib import colors

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_elements(num_pages)
            super().showPage()
        super().save()

    def draw_page_elements(self, page_count):
        self.saveState()
        
        # Suppress headers/footers on cover page (page 1)
        if self._pageNumber > 1:
            # Running Header
            self.setFont("Helvetica-Bold", 8)
            self.setFillColor(colors.HexColor("#1F4E79"))
            header_text = getattr(self, 'doc_header_title', 'BAAY PROJECTS LIMITED (RC 1526224)')
            self.drawString(36, self._pagesize[1] - 25, header_text)
            
            self.setFont("Helvetica-Oblique", 7.5)
            self.setFillColor(colors.HexColor("#595959"))
            sub_text = getattr(self, 'doc_header_sub', 'Audited Financial Statements & Statutory Working Papers')
            self.drawRightString(self._pagesize[0] - 36, self._pagesize[1] - 25, sub_text)
            
            self.setStrokeColor(colors.HexColor("#1F4E79"))
            self.setLineWidth(0.75)
            self.line(36, self._pagesize[1] - 28, self._pagesize[0] - 36, self._pagesize[1] - 28)
            
            # Running Footer
            self.setStrokeColor(colors.HexColor("#D9D9D9"))
            self.setLineWidth(0.5)
            self.line(36, 32, self._pagesize[0] - 36, 32)
            
            self.setFont("Helvetica", 7.5)
            self.setFillColor(colors.HexColor("#595959"))
            footer_left = getattr(self, 'doc_footer_left', 'BAAY PROJECTS LIMITED — IFRS REPORTING PACK')
            self.drawString(36, 22, footer_left)
            
            page_str = f"Page {self._pageNumber} of {page_count}"
            self.drawRightString(self._pagesize[0] - 36, 22, page_str)
            
        self.restoreState()
