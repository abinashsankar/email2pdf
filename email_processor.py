import os
from compoundfiles import CompoundFileReader
from email import policy
from email.parser import BytesParser
from datetime import datetime, timedelta
import io
import warnings

# Suppress all warnings from compoundfiles.streams (temporary workaround)
warnings.filterwarnings("ignore", category=Warning, module="compoundfiles.streams")

class EmailProcessor:
    """Processes .msg and .eml files to extract metadata and attachments."""
    
    def __init__(self, file_path, output_dir):
        self.file_path = file_path
        self.output_dir = output_dir
        self.email_data = {"From": None, "To": [], "Sent On": None, "CC": None, "Subject": None, "Body": None}
        self.attachments = []
        os.makedirs(output_dir, exist_ok=True)

    def process(self):
        """Process the email file based on its extension."""
        if self.file_path.lower().endswith('.msg'):
            self._process_msg()
        elif self.file_path.lower().endswith('.eml'):
            self._process_eml()
        return self.email_data, self.attachments

    def _parse_filetime(self, data):
        """Parse MAPI FILETIME to datetime."""
        try:
            filetime = int.from_bytes(data, 'little')
            epoch = datetime(1601, 1, 1)
            return epoch + timedelta(microseconds=filetime / 10)
        except Exception as e:
            print(f"Error parsing FILETIME: {e}")
            return "Unknown"

    def _process_msg(self):
        """Process .msg file using CompoundFileReader."""
        try:
            with CompoundFileReader(self.file_path) as doc:
                for entry in doc.root:
                    try:
                        self._extract_metadata(entry, doc)
                        self._extract_recipients(entry, doc)
                        self._extract_attachment(entry, doc)
                    except Exception as e:
                        print(f"Error processing entry {entry.name}: {e}")
        except Exception as e:
            print(f"Failed to process .msg file {self.file_path}: {e}")

    def _extract_metadata(self, entry, doc):
        """Extract metadata from .msg streams."""
        if not entry.isdir:
            try:
                with doc.open(entry) as stream:  # Use context manager
                    if stream is None:
                        print(f"Stream is None for entry: {entry.name}")
                        return
                    data = stream.read()
                    mappings = {
                        '__substg1.0_0C1A001F': ("From", lambda x: x.decode('utf-16-le', errors='ignore').rstrip('\x00')),
                        '__substg1.0_0037001F': ("Subject", lambda x: x.decode('utf-16-le', errors='ignore').rstrip('\x00')),
                        '__substg1.0_1000001F': ("Body", lambda x: x.decode('utf-16-le', errors='ignore').rstrip('\x00')),
                        '__substg1.0_0E03001F': ("CC", lambda x: x.decode('utf-16-le', errors='ignore').rstrip('\x00')),
                        '__substg1.0_00390040': ("Sent On", self._parse_filetime),
                        '__substg1.0_0E060040': ("Sent On", self._parse_filetime),
                    }
                    if entry.name in mappings and (mappings[entry.name][0] != "Sent On" or not self.email_data["Sent On"]):
                        key, func = mappings[entry.name]
                        self.email_data[key] = func(data)
            except Exception as e:
                print(f"Error extracting metadata from {entry.name}: {e}")

    def _extract_recipients(self, entry, doc):
        """Extract recipients from .msg recip_version directories."""
        if entry.isdir and entry.name.startswith('__recip_version1.0_'):
            for recip_stream in entry:
                if recip_stream.name == '__substg1.0_3003001F':
                    try:
                        with doc.open(recip_stream) as stream:  # Use context manager
                            if stream is None:
                                print(f"Stream is None for recipient: {recip_stream.name}")
                                continue
                            recipient = stream.read().decode('utf-16-le', errors='ignore').rstrip('\x00')
                            self.email_data["To"].append(recipient)
                    except Exception as e:
                        print(f"Error extracting recipient from {recip_stream.name}: {e}")

    def _extract_attachment(self, entry, doc):
        """Extract and save attachments from .msg."""
        if entry.isdir and entry.name.startswith('__attach_version1.0_'):
            attachment = {"data": None, "filename": None, "mime_type": None}
            for stream_entry in entry:
                if not stream_entry.isdir:
                    try:
                        with doc.open(stream_entry) as stream:  # Use context manager
                            if stream is None:
                                print(f"Stream is None for attachment: {stream_entry.name}")
                                continue
                            data = stream.read()
                            if stream_entry.name == '__substg1.0_37010102':
                                attachment["data"] = data
                            elif stream_entry.name in ('__substg1.0_370E001F', '__substg1.0_3707001F'):
                                attachment["filename"] = data.decode('utf-16-le', errors='ignore').rstrip('\x00')
                            elif stream_entry.name == '__substg1.0_3704001F':
                                attachment["mime_type"] = data.decode('utf-16-le', errors='ignore').rstrip('\x00')
                    except Exception as e:
                        print(f"Error extracting attachment from {stream_entry.name}: {e}")
            if attachment["data"]:
                try:
                    filename = self._save_attachment(attachment, entry.name)
                    self.attachments.append(filename)
                except Exception as e:
                    print(f"Error saving attachment for {entry.name}: {e}")

    def _process_eml(self):
        """Process .eml file using email.parser."""
        with open(self.file_path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        self.email_data["From"] = msg.get("From", "Unknown Sender")
        self.email_data["To"] = msg.get("To", "").split(", ") if msg.get("To") else []
        self.email_data["CC"] = msg.get("CC", "")
        self.email_data["Subject"] = msg.get("Subject", "No Subject")
        sent_on = msg.get("Date")
        if sent_on:
            try:
                self.email_data["Sent On"] = datetime.strptime(sent_on, "%a, %d %b %Y %H:%M:%S %z")
            except ValueError:
                self.email_data["Sent On"] = "Unknown"
        self._extract_eml_body(msg)
        self._extract_eml_attachments(msg)

    def _extract_eml_body(self, msg):
        """Extract body from .eml file."""
        body_found = False
        for part in msg.walk():
            if part.get_content_type().startswith("text/") and not part.get("Content-Disposition"):
                try:
                    body = part.get_payload(decode=True)
                    if body:
                        self.email_data["Body"] = body.decode("utf-8", errors="ignore").strip()
                        body_found = True
                        break
                except Exception as e:
                    print(f"Error decoding body: {e}")
        if not body_found:
            self.email_data["Body"] = "No body content found in .eml file."

    def _extract_eml_attachments(self, msg):
        """Extract attachments from .eml file."""
        for part in msg.walk():
            if part.get("Content-Disposition") and "attachment" in part.get("Content-Disposition"):
                filename = part.get_filename() or f"attachment_{len(self.attachments)}.bin"
                filepath = os.path.join(self.output_dir, filename)
                with open(filepath, "wb") as f:
                    payload = part.get_payload(decode=True)
                    if payload:
                        f.write(payload)
                self.attachments.append(filename)
                print(f"Saved attachment: {filepath}")

    def _save_attachment(self, attachment, entry_name):
        """Save attachment with proper filename."""
        mime_to_ext = {
            'application/pdf': '.pdf', 'text/plain': '.txt', 'image/jpeg': '.jpg',
            'image/png': '.png', 'application/msword': '.doc'
        }
        base_filename = attachment["filename"] or f"attachment_{entry_name[-8:]}"
        for char in '<>:"/\\|?*':
            base_filename = base_filename.replace(char, '.')
        ext = mime_to_ext.get(attachment["mime_type"], '.bin') if '.' not in base_filename else ''
        filename = base_filename + ext
        filepath = os.path.join(self.output_dir, filename)
        with open(filepath, "wb") as f:
            f.write(attachment["data"])
        print(f"Saved attachment: {filepath}")
        return filename