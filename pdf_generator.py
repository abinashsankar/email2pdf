from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
from textwrap import wrap

class PDFGenerator:
    """Generates a PDF from email data and attachments."""
    
    def __init__(self, email_data, attachments, output_path="email_content.pdf"):
        self.email_data = email_data
        self.attachments = attachments
        self.output_path = output_path
        self.width, self.height = letter  # 612 x 792 points
        self.margin = 30
        self.line_height = 14

    def generate(self):
        """Generate and save the PDF."""
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        y = self.height - self.margin - 20

        y = self._draw_headers(can, y)
        y -= 10
        y = self._draw_body(can, y)
        y -= 10
        y = self._draw_attachments(can, y)

        if not any(self.email_data.values()) and not self.attachments:
            self._draw_text(can, "No email content or attachments found.", self.margin, y)

        can.save()
        self._save_pdf(packet)

    def _draw_text(self, can, text, x, y, font="Helvetica", size=12, bold=False):
        """Draw text with wrapping and page breaks."""
        can.setFont(f"{font}{'-Bold' if bold else ''}", size)
        wrapped_lines = wrap(text, width=int((self.width - 2 * self.margin) / (size * 0.6)))
        for line in wrapped_lines:
            if y < self.margin + 20:
                can.showPage()
                can.setFont(f"{font}{'-Bold' if bold else ''}", size)
                y = self.height - self.margin - 20
            can.drawString(x, y, line)
            y -= self.line_height
        return y

    def _draw_headers(self, can, y):
        """Draw email headers."""
        label_width = 80
        for key in ["From", "To", "Sent On", "CC", "Subject"]:
            value = self.email_data.get(key)
            if value or key == "Sent On":
                label = f"{key}:"
                value_str = ", ".join(value) if key == "To" and value else str(value) if value else "Not Found"
                if key == "Subject" and len(value_str) > 50:
                    value_str = value_str[:47] + "..."
                can.setFont("Helvetica-Bold", 12)
                can.drawString(self.margin, y, label)
                can.setFont("Helvetica", 12)
                y = self._draw_text(can, value_str, self.margin + label_width, y)
                y -= 5
        return y

    def _draw_body(self, can, y):
        """Draw email body with chain detection."""
        body = self.email_data.get("Body", "No body content found.")
        if body == None:
            body = ""
        paragraphs = body.split('\n')
        in_chain = False
        for para in paragraphs:
            if para.startswith(">") or "-----Original Message-----" in para or "From:" in para:
                in_chain = True
                can.setFillColorRGB(0.5, 0.5, 0.5)
            elif in_chain and not para.strip():
                in_chain = False
                can.setFillColorRGB(0, 0, 0)
            y = self._draw_text(can, para, self.margin + (10 if in_chain else 0), y)
            if not para and y > self.margin:
                y -= self.line_height
        can.setFillColorRGB(0, 0, 0)
        return y

    def _draw_attachments(self, can, y):
        """Draw attachments list."""
        if self.attachments:
            y = self._draw_text(can, "Attachments:", self.margin, y, bold=True)
            for attach in self.attachments:
                y = self._draw_text(can, f"- {attach}", self.margin + 10, y)
        return y

    def _save_pdf(self, packet):
        """Save the PDF to disk."""
        packet.seek(0)
        pdf_reader = PdfReader(packet)
        pdf_writer = PdfWriter()
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
        with open(self.output_path, "wb") as f:
            pdf_writer.write(f)
        print(f"Saved PDF to {self.output_path} with {len(pdf_reader.pages)} pages")