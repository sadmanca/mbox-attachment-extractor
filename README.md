# mbox Attachment Extractor

## Overview

This Python script (`extract.py`) is designed to parse `.mbox` files containing email messages, extract attachments from those emails, and optionally save each email as an individual `.mbox` file in corresponding folders. It provides a command-line interface (CLI) to customize its behavior based on various options.

## Features

- **Attachment Extraction**: Automatically detects and saves attachments from emails in the `.mbox` file.
- **Individual Email Saving**: Optionally saves each email as a separate `.mbox` file in a directory named after the email's timestamp and subject.
- **Logging**: Logs detailed information about the processing steps, including saved emails, skipped emails, and errors.
- **Progress Monitoring**: Uses `tqdm` to display a progress bar in the terminal, indicating the status of email processing.
- **Customizable Output**: Allows specifying an output folder where attachments and `.mbox` files are saved.

## Requirements

- Python 3.x
- Required Python packages (`colorama`, `tqdm`)


## Usage

### Command-Line Options

The script accepts several command-line options to customize its behavior:

- **Positional Arguments**:
- `mbox_file`: Path to the `.mbox` file to be processed.

- **Optional Arguments**:
- `-s, --save-mbox`: Save each email as an individual `.mbox` file in folders. Default is `False`.
- `-v, --verbose`: Print detailed logs to the terminal in addition to saving them to a log file. Default is `False`.
- `-l, --log-first-line`: Include the first line of each email (From field) in the log file. Default is `False`.
- `-o OUTPUT_FOLDER, --output-folder OUTPUT_FOLDER`: Specify the output folder where attachments and `.mbox` files will be saved. Default is `emails`.

### Examples

1. **Basic Usage**:

- Processes `mboxfile.mbox` without saving individual `.mbox` files, without verbose logging, and saving output in the default `emails` folder.

```
python extract.py path/to/your/mboxfile.mbox
```  

2. **Save Individual `.mbox` Files**:

- Processes `mboxfile.mbox`, saves each email as an individual `.mbox` file in default `emails` folder.

```
python extract.py path/to/your/mboxfile.mbox -s
```

3. **Verbose Output**:

- Processes `mboxfile.mbox`, prints detailed logs to the terminal, and saves output in the default `emails` folder.

```
python extract.py path/to/your/mboxfile.mbox -v
```

4. **Custom Output Folder**:

- Processes `mboxfile.mbox`, saves attachments and `.mbox` files in `custom_output_folder`.

```
python extract.py path/to/your/mboxfile.mbox -o "custom_output_folder/path""
```


5. **Verbose and Log First Line**:

- Processes `mboxfile.mbox`, prints detailed logs to the terminal with the first line of each email included, and saves output in the default `emails` folder.

```
python extract.py path/to/your/mboxfile.mbox -v -l
```


### Output

- **Attachments**: Extracted attachments from emails are saved in subfolders within the specified or default output folder.
- **Log File**: Detailed logs of the processing steps are saved in a timestamped log file (`email_parser_<timestamp>.log`).

```
python extract.py -h
```


### Notes

- Ensure Python and required packages (`colorama`, `tqdm`) are installed before running the script.
- Use `-h` or `--help` for help with command-line options:

