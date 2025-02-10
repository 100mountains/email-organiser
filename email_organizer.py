import os
import re
import shutil
import logging
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, List, Set, Tuple
from dataclasses import dataclass
import time
from email.utils import parseaddr, parsedate_to_datetime
from email import policy, message_from_file
from tenacity import retry, stop_after_attempt, wait_exponential

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('email_organizer.log')]
)
logger = logging.getLogger(__name__)

@dataclass
class EmailDomains:
    from_domain: str
    to_domains: List[str]
    cc_domains: List[str]
    gov_domains: Set[str]

@dataclass
class EmailDateInfo:
    datetime: datetime
    date_str: str
    year: str

class EmailProgress:
    def __init__(self, total_files):
        self.total = total_files
        self.processed = 0
        self.gov_emails = 0
        self.current_file = ""
        self.start_time = time.time()
        self.attachments_found = 0
        self.last_attachment = ""
        print("\n\n\n")  # Make initial space for three lines of progress
        
    def update(self, filepath, is_gov=False, attachment_copied: Optional[str] = None, increment_processed=True):
        if increment_processed:  # Only increment for original files, not copies
            self.processed += 1
        if is_gov:
            self.gov_emails += 1
        self.current_file = os.path.basename(filepath)
        if attachment_copied:
            self.attachments_found += 1
            self.last_attachment = os.path.basename(attachment_copied)
        self._display_progress()
        
    def _display_progress(self):
        width = 50
        percent = min(100, (self.processed / self.total) * 100)  # Cap at 100%
        filled = int((width * percent) / 100)
        bar = '=' * filled + '>' + ' ' * (width - filled - 1)
        elapsed = time.time() - self.start_time
        rate = self.processed / elapsed if elapsed > 0 else 0
        
        # Clear three lines
        print('\033[2K\033[1G', end='')  # Clear current line
        print('\033[1A\033[2K\033[1G', end='')  # Move up and clear line
        print('\033[1A\033[2K\033[1G', end='')  # Move up and clear line
        
        # Print all three lines
        if self.last_attachment:
            print(f"Last attachment: {self.last_attachment[:60]}... (Total: {self.attachments_found})")
        else:
            print(f"Attachments found: {self.attachments_found}")
        print(f"Processing: [{bar}] {percent:.1f}% ({self.processed}/{self.total}) {rate:.1f} files/s")
        print(f"Current: {self.current_file[:60]}... Found {self.gov_emails} gov.uk emails", end='\r')

def extract_email_details(file_path: str) -> Dict:
    headers = {
        'From': '',
        'To': '',
        'CC': '',  # Initialize CC as empty string
        'Subject': '',
        'Date': ''
    }
    
    encodings = ['utf-8', 'cp1252', 'iso-8859-1', 'windows-1254']
    content = None
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
            break
        except UnicodeDecodeError:
            continue
    
    if content is None:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

    if not looks_like_email(content):
        return headers

    soup = BeautifulSoup(content, 'html.parser')
    
    for header in headers.keys():
        value = None
        for tag in ['span', 'div', 'td', 'th', 'p']:
            element = (
                soup.find(tag, string=lambda x: x and f'{header}:' in x) or
                soup.find(tag, string=lambda x: x and header.lower() in x.lower()) or
                soup.find(tag, attrs={'class': lambda x: x and header.lower() in x.lower()})
            )
            if element:
                if element.next_sibling and isinstance(element.next_sibling, str):
                    value = element.next_sibling.strip()
                elif element.parent and element.parent.next_sibling and isinstance(element.parent.next_sibling, str):
                    value = element.parent.next_sibling.strip()
                elif ':' in element.text:
                    value = element.text.split(':', 1)[1].strip()
                break
        
        if value:
            headers[header] = value

    if not headers.get('Date'):
        date_patterns = [
            r'Date:\s*</div>([^<]+)',
            r'Date:\s*(\d{1,2}/\d{1,2}/\d{4},?\s*\d{1,2}:\d{2})',
            r'Date:\s*(\d{4}-\d{2}-\d{2}\s*\d{1,2}:\d{2})',
            r'Sent:\s*(\d{1,2}/\d{1,2}/\d{4})',
            r'Sent:\s*(\d{4}-\d{2}-\d{2})'
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, content)
            if match:
                headers['Date'] = match.group(1).strip()
                break

    return headers

def extract_eml_details(eml_file: str) -> Dict:
    headers = {
        'From': '',
        'To': '',
        'CC': '',
        'Subject': '',
        'Date': ''
    }

    # If it's a file named exactly '.eml' in an Attachments folder, read its target link
    if os.path.basename(eml_file) == '.eml' and 'Attachments-' in eml_file:
        try:
            # These files often contain a pointer to the actual email
            with open(eml_file, 'r', encoding='utf-8', errors='ignore') as f:
                possible_link = f.read().strip()
                if possible_link and os.path.exists(possible_link):
                    eml_file = possible_link  # Use the linked file instead
        except Exception as e:
            logger.debug(f"Failed to read link from .eml placeholder: {str(e)}")
    
    try:
        with open(eml_file, 'rb') as f:
            msg = message_from_file(f, policy=policy.default)
            
        for header in headers.keys():
            headers[header] = str(msg[header]) if msg[header] else ''
            
        if any(headers.values()):  # If we got any values, return them
            return headers
            
    except Exception as e:
        logger.debug(f"Standard EML parsing failed for {eml_file}: {str(e)}")
    
    try:
        with open(eml_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            
        if looks_like_email(content):
            return extract_email_details(eml_file)
            
    except Exception as e:
        logger.debug(f"HTML fallback parsing failed for {eml_file}: {str(e)}")
    
    try:
        with open(eml_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            
        header_patterns = {
            'From': r'From:\s*([^\n]+)',
            'To': r'To:\s*([^\n]+)',
            'CC': r'CC:\s*([^\n]+)',
            'Subject': r'Subject:\s*([^\n]+)',
            'Date': r'Date:\s*([^\n]+)'
        }
        
        for header, pattern in header_patterns.items():
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                headers[header] = match.group(1).strip()
                
    except Exception as e:
        logger.error(f"All parsing methods failed for {eml_file}: {str(e)}")
    
    return headers

def looks_like_email(content: str) -> bool:
    email_indicators = [
        'From:', 'To:', 'Sent:', 'Date:', 'Subject:',
        'mailto:', '@', 'Reply-To:', 'Cc:', 'Bcc:'
    ]
    content_lower = content.lower()
    matches = sum(1 for indicator in email_indicators if indicator.lower() in content_lower)
    return matches >= 3

def get_domain(email_str: str) -> str:
    name, addr = parseaddr(email_str)
    if '@' in addr:
        return addr.split('@')[1].strip().lower()
    return 'unknown'

def create_email_path(base_dir: str, domain: str, year: str, subject: str) -> str:
    safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)
    safe_subject = safe_subject[:100]  # Limit subject length
    path = os.path.join(base_dir, domain, year, safe_subject)
    os.makedirs(path, exist_ok=True)
    return path

def copy_with_metadata(src_path: str, dst_path: str, internal_dt: Optional[datetime] = None):
    try:
        shutil.copy2(src_path, dst_path)
        if internal_dt:
            new_time = internal_dt.timestamp()
            os.utime(dst_path, (new_time, new_time))
        else:
            st = os.stat(src_path)
            os.utime(dst_path, (st.st_atime, st.st_mtime))
    except Exception as e:
        logger.error(f"Failed to copy {src_path} to {dst_path}: {str(e)}")
        raise

def extract_embedded_attachments(eml_file: str, dst_dir: str) -> List[str]:
    """Extract attachments embedded within an EML file"""
    copied_attachments = []
    
    try:
        with open(eml_file, 'rb') as f:
            msg = message_from_file(f, policy=policy.default)
            
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
                
            # Skip the email body itself
            if part.get_content_maintype() == 'text' and part.get_content_disposition() is None:
                continue

            filename = part.get_filename()
            if filename:
                # Clean filename and ensure uniqueness
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                dst_file = os.path.join(dst_dir, safe_filename)
                
                # Handle duplicate filenames
                counter = 1
                while os.path.exists(dst_file):
                    base, ext = os.path.splitext(safe_filename)
                    dst_file = os.path.join(dst_dir, f"{base}_{counter}{ext}")
                    counter += 1

                # Extract and save the attachment
                try:
                    with open(dst_file, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    copied_attachments.append(dst_file)
                    logger.info(f"Extracted embedded attachment: {filename} -> {dst_file}")
                except Exception as e:
                    logger.error(f"Failed to save embedded attachment {filename}: {str(e)}")

    except Exception as e:
        logger.error(f"Error extracting embedded attachments from {eml_file}: {str(e)}")
        
    return copied_attachments

def copy_attachments(src_dir: str, dst_dir: str, email_file: str) -> List[str]:
    copied_attachments = []

    # Handle linked attachments from HTML files
    if email_file.endswith('.html'):
        try:
            with open(email_file, 'r', encoding='utf-8', errors='ignore') as f:
                soup = BeautifulSoup(f.read(), 'html.parser')
                
            attachment_links = soup.find_all('a', href=True)
            src_path = os.path.dirname(email_file)
            
            for link in attachment_links:
                path = link.get('href')
                
                if not path or path.startswith(('mailto:', 'http:', 'https:', 'tel:', '#', 'data:')):
                    continue
                    
                if path.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp')):
                    continue
                    
                if os.path.isabs(path):
                    src_file = os.path.join(src_path, os.path.basename(path))
                else:
                    src_file = os.path.normpath(os.path.join(src_path, path))
                
                if not src_file.startswith(src_path):
                    logger.warning(f"Skipping potentially unsafe path: {src_file}")
                    continue
                    
                dst_file = os.path.join(dst_dir, os.path.basename(src_file))
                
                if os.path.exists(src_file) and os.path.isfile(src_file):
                    copy_with_metadata(src_file, dst_file)
                    copied_attachments.append(dst_file)
                    logger.info(f"Copied linked attachment: {src_file} -> {dst_file}")

        except Exception as e:
            logger.error(f"Error processing HTML attachments for {email_file}: {str(e)}")

    # Handle EML files - both linked and embedded attachments
    if email_file.endswith('.eml'):
        # First try to handle any linked attachments if this is a placeholder
        if os.path.basename(email_file) == '.eml' and 'Attachments-' in email_file:
            try:
                with open(email_file, 'r', encoding='utf-8', errors='ignore') as f:
                    possible_link = f.read().strip()
                    if possible_link and os.path.exists(possible_link):
                        # Copy linked attachments
                        src_path = os.path.dirname(possible_link)
                        for item in os.listdir(src_path):
                            if os.path.isfile(os.path.join(src_path, item)):
                                src_file = os.path.join(src_path, item)
                                dst_file = os.path.join(dst_dir, item)
                                copy_with_metadata(src_file, dst_file)
                                copied_attachments.append(dst_file)
                                logger.info(f"Copied linked attachment: {src_file} -> {dst_file}")
            except Exception as e:
                logger.error(f"Error processing linked attachments for {email_file}: {str(e)}")

        # Then extract any embedded attachments
        embedded_attachments = extract_embedded_attachments(email_file, dst_dir)
        copied_attachments.extend(embedded_attachments)
            
    return copied_attachments

class EmailProcessor:
    def __init__(self, root_dir: str, output_dir: str):
        self.root_dir = Path(root_dir)
        self.output_dir = Path(output_dir)
        self.date_parser = EmailDateParser()
        self.logger = logging.getLogger(__name__)
    
    def process_emails(self):
        """Main processing function"""
        email_files = self.find_email_files()
        logger.info(f"Found {len(email_files)} potential email files to process")
        
        progress = EmailProgress(len(email_files))
        
        for file_path in email_files:
            self.process_single_email(file_path, progress)
            
        print("\nProcessing complete!")
        logger.info("Processing complete!")
    
    def find_email_files(self) -> list[Path]:
        """Find all email files recursively"""
        email_files = []
        # Check for normal extensions
        for ext in ['.html', '.eml']:
            email_files.extend([
                f for f in self.root_dir.rglob(f'*{ext}')
                if f.name.lower() != 'index.html'
            ])
        # Also check for literal '.eml' files
        email_files.extend(list(self.root_dir.rglob('.eml')))
        return email_files

    def process_single_email(self, file_path: Path, progress: EmailProgress) -> None:
        """Process a single email file"""
        try:
            # If it's a file named exactly '.eml' in an Attachments folder, skip extraction if empty
            if (file_path.name == '.eml' and 
                'Attachments-' in str(file_path.parent) and 
                file_path.stat().st_size == 0):
                logger.debug(f"Skipping empty attachment placeholder: {file_path}")
                progress.update(str(file_path))
                return

            details = self.extract_email_details(file_path)
            if not details:
                self.logger.warning(f"Could not extract details from {file_path}")
                progress.update(str(file_path))
                return
                
            domains = self._extract_domains(details)
            if not domains.gov_domains:
                self.logger.debug(f"Skipping {file_path} - no .gov.uk domains found")
                progress.update(str(file_path))
                return
                
            date_info = self._get_date_info(details, file_path)
            self._process_primary_copy(file_path, details, domains, date_info, progress)
            self._process_gov_copies(file_path, details, domains, date_info, progress)
            
        except Exception as e:
            self.logger.error(f"Error processing {file_path}: {str(e)}", exc_info=True)
            progress.update(str(file_path))
    
    def extract_email_details(self, file_path: Path) -> Optional[Dict]:
        """Extract details from either HTML or EML file"""
        if file_path.suffix == '.html':
            return extract_email_details(str(file_path))
        elif file_path.suffix == '.eml' or file_path.name == '.eml':
            return extract_eml_details(str(file_path))
        return None
    
    def _extract_domains(self, details: Dict) -> EmailDomains:
        """Extract and categorize all domains from email details"""
        from_domain = get_domain(details.get('From', ''))
        to_domains = [get_domain(addr.strip()) 
                     for addr in details.get('To', '').split(',') 
                     if addr.strip()]
        cc_domains = [get_domain(addr.strip()) 
                     for addr in details.get('CC', '').split(',') 
                     if addr.strip()]
                     
        all_domains = [from_domain] + to_domains + cc_domains
        gov_domains = [d for d in all_domains if d and d.endswith('.gov.uk')]
        
        return EmailDomains(
            from_domain=from_domain or 'unknown',
            to_domains=to_domains,
            cc_domains=cc_domains,
            gov_domains=set(gov_domains)
        )
    
    def _get_date_info(self, details: Dict, file_path: Path) -> EmailDateInfo:
        """Extract and parse date information"""
        internal_dt = self.date_parser.parse_date(
            details.get('Date'),
            filename=file_path.name
        ) or datetime.now()
        
        return EmailDateInfo(
            datetime=internal_dt,
            date_str=self.date_parser.format_date(internal_dt),
            year=internal_dt.strftime('%Y')
        )
    
    def _process_primary_copy(self, file_path: Path, details: Dict, 
                            domains: EmailDomains, date_info: EmailDateInfo,
                            progress: EmailProgress) -> None:
        """Process the primary copy of the email"""
        primary_path = create_email_path(
            str(self.output_dir), 
            domains.from_domain,
            date_info.year, 
            details.get('Subject', 'No Subject')
        )
        
        dst_file = Path(primary_path) / f"{date_info.date_str}_{file_path.name}"
        self._safe_copy(file_path, dst_file, date_info.datetime)
        self._copy_attachments(file_path, dst_file, progress)
        progress.update(str(file_path), bool(domains.gov_domains))
    
    def _process_gov_copies(self, file_path: Path, details: Dict,
                          domains: EmailDomains, date_info: EmailDateInfo,
                          progress: EmailProgress) -> None:
        """Process copies for government domains"""
        for domain in sorted(domains.gov_domains - {domains.from_domain}):
            path = create_email_path(
                str(self.output_dir),
                domain,
                date_info.year,
                details.get('Subject', 'No Subject')
            )
            
            dst_file = Path(path) / f"{date_info.date_str}_{file_path.name}"
            self._safe_copy(file_path, dst_file, date_info.datetime)
            self._copy_attachments(file_path, dst_file, progress)
            progress.update(str(file_path), True, increment_processed=False)
    
    def _safe_copy(self, src: Path, dst: Path, timestamp: datetime) -> None:
        """Safely copy a file with proper error handling"""
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            copy_with_metadata(str(src), str(dst), timestamp)
        except Exception as e:
            self.logger.error(f"Failed to copy {src} to {dst}: {str(e)}")
            raise
    
    def _copy_attachments(self, src: Path, dst: Path, progress: EmailProgress) -> None:
        """Copy attachments with proper error handling"""
        try:
            copied = copy_attachments(str(src.parent), str(dst.parent), str(src))
            for attachment in copied:
                progress.update(str(src), attachment_copied=attachment, increment_processed=False)
        except Exception as e:
            self.logger.error(f"Failed to copy attachments for {src}: {str(e)}")

class EmailDateParser:
    """A robust email date parser with standardized formats"""
    
    DATE_FORMATS = [
        '%d/%m/%Y, %H:%M',     # 11/10/2024, 15:13
        '%d/%m/%Y %H:%M',      # 11/10/2024 15:13
        '%Y-%m-%d',            # 2024-02-10
        '%d/%m/%Y',            # 10/02/2024
        '%Y-%m-%d %H:%M',      # 2024-02-10 15:30
        '%Y-%m-%dT%H:%M:%S',   # 2024-02-10T15:30:00
        '%a, %d %b %Y %H:%M:%S %z'  # Sun, 10 Feb 2024 15:30:00 +0000
    ]

    @classmethod
    def parse_date(cls, date_str: Optional[str], filename: Optional[str] = None) -> Optional[datetime]:
        if not date_str and not filename:
            return None
            
        if date_str:
            try:
                return parsedate_to_datetime(date_str.strip())
            except (TypeError, ValueError):
                pass
                
            for fmt in cls.DATE_FORMATS:
                try:
                    return datetime.strptime(date_str.strip(), fmt)
                except ValueError:
                    continue
                    
            date_only = cls._extract_date_portion(date_str)
            if date_only:
                try:
                    return datetime.strptime(date_only, '%Y-%m-%d')
                except ValueError:
                    pass

        if filename:
            patterns = [
                r'[-_](\d{8})[-_]',       # matches -20240210- or _20240210_
                r'(\d{8})[._]',           # matches 20240210. or 20240210_
                r'[-_](\d{6})[-_]',       # matches -240210- or _240210_
                r'(\d{4}-\d{2}-\d{2})',   # matches 2024-02-10
                r'(\d{2}-\d{2}-\d{4})',   # matches 10-02-2024
            ]
            
            for pattern in patterns:
                match = re.search(pattern, filename)
                if match:
                    date_str = match.group(1)
                    try:
                        if len(date_str) == 8:
                            return datetime.strptime(date_str, '%Y%m%d')
                        elif len(date_str) == 6:
                            return datetime.strptime(f"20{date_str}", '%Y%m%d')
                        elif '-' in date_str:
                            if date_str.startswith('20'):
                                return datetime.strptime(date_str, '%Y-%m-%d')
                            else:
                                return datetime.strptime(date_str, '%d-%m-%Y')
                    except ValueError:
                        continue
        
        logger.warning(f"Failed to parse date from string: {date_str} or filename: {filename}")
        return datetime.now()

    @staticmethod
    def _extract_date_portion(date_str: str) -> Optional[str]:
        patterns = [
            r'(\d{4}-\d{2}-\d{2})',
            r'(\d{2}/\d{2}/\d{4})',
            r'(\d{4}/\d{2}/\d{2})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, date_str)
            if match:
                return match.group(1)
                
        return None

    @classmethod
    def format_date(cls, dt: datetime) -> str:
        return dt.strftime('%Y%m%d')

def main():
    input_dir = './EMAIL-MAIN'  # Adjust as needed
    output_dir = './sorted_emails'
    
    os.makedirs(output_dir, exist_ok=True)
    logger.info(f"Starting email organization from {input_dir} to {output_dir}")
    
    try:
        processor = EmailProcessor(input_dir, output_dir)
        processor.process_emails()
        logger.info("Email organization completed successfully")
    except Exception as e:
        logger.error(f"Failed to complete email organization: {str(e)}", exc_info=True)
        raise

if __name__ == '__main__':
    main()