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
custom_tqdm_format = f"{GREEN}{{desc}} {{percentage:3.0f}}%|{{bar}}| {{n_fmt}}/{{total_fmt}} [{{elapsed}}<{{remaining}}]{RESET}"

# Invalid characters for Windows and Linux directory names
INVALID_CHARS = r'[<>:"/\\|?*]'

def sanitize_folder_name(name):
    return re.sub(INVALID_CHARS, '_', name)

def is_image_attachment(part):
    return part.get_content_maintype() == 'image'

def save_attachment(msg, download_folder, extract_all_attachments):
    saved = False
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        if not extract_all_attachments and not is_image_attachment(part):
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
    if not logger.handlers:  # Check if handlers already exist
        logger.setLevel(logging.DEBUG)
        fh = logging.FileHandler(log_filename)
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)

def log_email_info(index, folder_name, from_line, log_first_line, is_saved, verbose, log_skipped):
    if is_saved:
        status_message = f"Saved email {index+1} to folder: {folder_name}"
        color = Fore.GREEN
    elif log_skipped:
        status_message = f"Skipped email {index+1} (no attachments)"
        color = Fore.LIGHTBLACK_EX
    else:
        return
    
    if log_first_line:
        status_message += f" | From {from_line.strip()}"
    
    message = f"{color}{status_message}{Style.RESET_ALL}"
    
    if verbose:
        print(message)  # Print the status message with color
    logging.info(status_message)  # Log the status message without color

def process_single_mbox(mbox_file, save_mbox, verbose, log_first_line, output_folder, display_full_path, extract_all_attachments, log_skipped):
    display_name = mbox_file if display_full_path else os.path.basename(mbox_file)
    logging.info(f"Processing mbox file: {mbox_file}")
    logging.info(f"Output folder: {output_folder}")
    
    mbox = mailbox.mbox(mbox_file)
    total_emails = len(mbox)
    with tqdm(total=total_emails, desc=display_name, bar_format=custom_tqdm_format) as pbar:
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
                    saved = save_attachment(email_message, folder_path, extract_all_attachments)
                    log_email_info(i, folder_name, from_line, log_first_line, saved, verbose, log_skipped)
                else:
                    log_email_info(i, None, from_line, log_first_line, False, verbose, log_skipped)
            pbar.update(1)
    print()  # Add a newline after the progress bar
    logging.info("Finished processing mbox file.")

def process_mbox_files_in_folder(input_folder, save_mbox, verbose, log_first_line, output_folder, display_full_path, extract_all_attachments, log_skipped):
    setup_logging(verbose, log_first_line)
    mbox_files = [f for f in os.listdir(input_folder) if f.endswith('.mbox')]
    
    for mbox_file in mbox_files:
        mbox_file_path = os.path.join(input_folder, mbox_file)
        mbox_output_folder = os.path.join(output_folder, os.path.splitext(mbox_file)[0])
        os.makedirs(mbox_output_folder, exist_ok=True)
        logging.info(f"Processing mbox file: {mbox_file_path}")
        logging.info(f"Output folder: {mbox_output_folder}")
        process_single_mbox(mbox_file_path, save_mbox, verbose, log_first_line, mbox_output_folder, display_full_path, extract_all_attachments, log_skipped)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Parse .mbox file and save emails with attachments.")
    parser.add_argument('path', type=str, help="Path to the .mbox file or directory containing .mbox files")
    parser.add_argument('-s', '--save-mbox', action='store_true', help="Save individual emails as .mbox files in folders")
    parser.add_argument('-v', '--verbose', action='store_true', help="Print logs to terminal in addition to saving them to a log file")
    parser.add_argument('-f', '--log-first-line', action='store_true', help="Include the first line of the email (From field) in the log")
    parser.add_argument('-o', '--output-folder', type=str, default='emails', help="Output folder name/path")
    parser.add_argument('-d', '--display-full-path', action='store_true', help="Display full file path instead of just the file name")
    parser.add_argument('-a', '--extract-all-attachments', action='store_true', help="Extract all types of attachments instead of just images")
    parser.add_argument('-l', '--log-skipped', action='store_true', help="Log skipped emails")
    args = parser.parse_args()

    custom_tqdm_format += "\n" if args.verbose else ""

    if os.path.isdir(args.path):
        process_mbox_files_in_folder(args.path, args.save_mbox, args.verbose, args.log_first_line, args.output_folder, args.display_full_path, args.extract_all_attachments, args.log_skipped)
    else:
        process_single_mbox(args.path, args.save_mbox, args.verbose, args.log_first_line, args.output_folder, args.display_full_path, args.extract_all_attachments, args.log_skipped)
