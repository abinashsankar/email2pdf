import os
from email_processor import EmailProcessor
from pdf_generator import PDFGenerator

def main():
    """Main function to process email files and generate PDFs."""
    msg_file = r"strangeDate.msg"  # Example path
    output_dir = r"attachments" # output Folder
    pdf_output = os.path.join(output_dir, "email_content.pdf")

    processor = EmailProcessor(msg_file, output_dir)
    email_data, attachments = processor.process()

    generator = PDFGenerator(email_data, attachments, pdf_output)
    generator.generate()

if __name__ == "__main__":
    main()