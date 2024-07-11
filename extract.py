import os
import argparse
import mailbox
from email import policy
from email.parser import BytesParser
from email.generator import BytesGenerator
from email.utils import parsedate_to_datetime
from tqdm import tqdm
from colorama import init, Fore, Style
import re
from io import BytesIO
import logging
from datetime import datetime

# Initialize colorama
init(autoreset=True)

GREEN = "\033[92m"
RESET = "\033[0m"
custom_tqdm_format = f"{GREEN}{{desc}} {{percentage:3.0f}}%|{{bar}}| {{n_fmt}}/{{total_fmt}} [{{elapsed}}<{{remaining}}{RESET}]"

# Invalid characters for Windows and Linux directory names
INVALID_CHARS = r'[<>:"/\\|?*]'

def sanitize_folder_name(name):
    return re.sub(INVALID_CHARS, '_', name)

def save_attachment(msg, download_folder):
    saved = False
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if filename:
            filepath = os.path.join(download_folder, filename)
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            saved = True
    return saved

def create_folder_name(msg):
    date_tuple = parsedate_to_datetime(msg['Date'])
    date_str = date_tuple.strftime("%Y-%m-%d_%H-%M-%S")
    subject = msg['Subject'] if msg['Subject'] else 'No_Subject'
    sanitized_subject = sanitize_folder_name(subject)
    return f"{date_str}_{sanitized_subject}"

def save_email_to_mbox(from_line, msg, folder_path):
    mbox_path = os.path.join(folder_path, 'email.mbox')
    with open(mbox_path, 'wb') as f:
        # Write the From line first
        f.write('From '.encode('utf-8'))
        f.write(from_line.encode('utf-8'))
        f.write(b'\n')
        # Use BytesIO to write the email message properly
        with BytesIO() as bio:
            gen = BytesGenerator(bio, policy=policy.default)
            gen.flatten(msg)
            f.write(bio.getvalue())

def setup_logging(verbose, log_first_line):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f'email_parser_{timestamp}.log'
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(log_filename)
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    if verbose:
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
    return log_first_line

def log_email_info(index, folder_name, from_line, log_first_line, is_saved, verbose):
    if is_saved:
        status_message = f"Saved email {index+1} to folder: {folder_name}"
        color = Fore.GREEN
    else:
        status_message = f"Skipped email {index+1} (no attachments)"
        color = Fore.LIGHTBLACK_EX
    
    if log_first_line:
        status_message += f" | From {from_line.strip()}"
    
    message = f"{color}{status_message}{Style.RESET_ALL}"
    logging.info(message)
    if verbose:
        print(message)

def parse_mbox(mbox_file, save_mbox, verbose, log_first_line, output_folder):
    setup_logging(verbose, log_first_line)
    logging.info(f"Processing mbox file: {mbox_file}")
    logging.info(f"Output folder: {output_folder}")
    
    mbox = mailbox.mbox(mbox_file)
    total_emails = len(mbox)
    with tqdm(total=total_emails, desc="Processing emails", bar_format=custom_tqdm_format) as pbar:
        for i, msg in enumerate(mbox):
            if isinstance(msg, mailbox.mboxMessage):
                email_message = BytesParser(policy=policy.default).parsebytes(msg.as_bytes())
                from_line = msg.get_from()
                if any(part.get_filename() for part in email_message.walk()):
                    folder_name = create_folder_name(email_message)
                    folder_path = os.path.join(output_folder, folder_name)
                    os.makedirs(folder_path, exist_ok=True)
                    if save_mbox:
                        save_email_to_mbox(from_line, email_message, folder_path)
                    saved = save_attachment(email_message, folder_path)
                    log_email_info(i, folder_name, from_line, log_first_line, saved, verbose)
                else:
                    log_email_info(i, None, from_line, log_first_line, False, verbose)
            pbar.update(1)
    logging.info("Finished processing mbox file.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Parse .mbox file and save emails with attachments.")
    parser.add_argument('mbox_file', type=str, help="Path to the .mbox file")
    parser.add_argument('-s', '--save-mbox', action='store_true', help="Save individual emails as .mbox files in folders")
    parser.add_argument('-v', '--verbose', action='store_true', help="Print logs to terminal in addition to saving them to a log file")
    parser.add_argument('-l', '--log-first-line', action='store_true', help="Include the first line of the email (From field) in the log")
    parser.add_argument('-o', '--output-folder', type=str, default='emails', help="Output folder name/path")
    args = parser.parse_args()
    
    parse_mbox(args.mbox_file, args.save_mbox, args.verbose, args.log_first_line, args.output_folder)
