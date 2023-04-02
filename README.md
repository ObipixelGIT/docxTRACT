# docxTRACT
docxTRACT extracts artifacts from a DOCX (Microsoft Word) file for OSINT and Digital Forensics analysis.

## How the script works?

Extracts files from a DOCX (Microsoft Word) file.
It uses the zipfile module to open the DOCX file and extract the files within.

The script defines a function called "extract_docx" which takes the following arguments:

- docx_file_path: the path to the DOCX file to extract from
- extract_images_only: a boolean flag indicating whether to only extract image files
- display_sha1: a boolean flag indicating whether to display the SHA1 hash of each extracted file
- destination_dir: the directory to extract files to (if not provided, files are extracted to the current directory)

The function returns a list of the extracted files.

The script also defines a "sha1sum" function that calculates the SHA1 hash of a file. This is used in the "extract_docx" function to display the SHA1 hash of each extracted file (if the "display_sha1" flag is set to True).

The script uses the argparse module to parse command line arguments. The available options are:

- filename: the path to the DOCX file to extract from (required)
- -x or --xtract: a flag indicating whether to extract files
- -i or --img_only: a flag indicating whether to only extract image files
- -s or --sha1: a flag indicating whether to display the SHA1 hash of each extracted file
- -d or --destdir: the directory to extract files to

If the "-x" or "--xtract" flag is not provided, the script will print the help message and exit.

If the specified DOCX file does not exist, the script will print an error message and exit.

When the script is run, it prints a message indicating any artifacts found in the DOCX file (i.e. the files that will be extracted).

Then, if the "-x" or "--xtract" flag is provided, it extracts the files and either displays the SHA1 hash of each extracted file (if the "-s" or "--sha1" flag is provided) or writes the SHA1 hashes to a file called "hashes.txt" in the destination directory (if the "-d" or "--destdir" flag is provided).

If the "-d" or "--destdir" flag is not provided, files are extracted to the current directory.

Finally, if files are extracted, the script prints a message indicating the directory they were extracted to.

## Preparation

The following Python modules must be installed:
```bash
pip3 install argparse, hashlib, zipfile
```

## Permissions

Ensure you give the script permissions to execute. Do the following from the terminal:
```bash
sudo chmod +x docxTRACT.py
```

## Usage
```bash
sudo python3 docxTRACT.py -h
usage: docxTRACT [-h] [-x] [-i] [-s] [-d DESTDIR] filename

docxTRACT version 0.1

positional arguments:
  filename              the DOCX file to extract files from

options:
  -h, --help            show this help message and exit
  -x, --xtract          extract files (by default to current dir)
  -i, --img_only        extract only images
  -s, --sha1            display extracted file SHA1 hash
  -d DESTDIR, --destdir DESTDIR
                        define destination directory

docxTRACT comes with ABSOLUTELY NO WARRANTY!
```

## Sample script
```python
#!/usr/bin/env python3

import argparse
import hashlib
import os
import zipfile

VERSION = "0.1"

def sha1sum(file_path):
    """Calculate SHA1 hash of a file."""
    with open(file_path, "rb") as file:
        return hashlib.sha1(file.read()).hexdigest()

def extract_docx(docx_file_path, extract_images_only=False, display_sha1=False, destination_dir=None):
    """Extract files from a DOCX file."""
    extracted_files = []

    with zipfile.ZipFile(docx_file_path) as docx:
        print()
        print('\033[1;41m ARTIFACTS FOUND: \033[m')
        print()

        for file in docx.infolist():
            file_path = file.filename

            if extract_images_only and not file_path.startswith("word/media/"):
                continue

            if destination_dir is not None:
                file_path = os.path.join(destination_dir, os.path.basename(file_path))

            if file_path.endswith("/"):
                os.makedirs(file_path, exist_ok=True)
            else:
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                with open(file_path, "wb") as output_file, docx.open(file) as input_file:
                    output_file.write(input_file.read())

                extracted_files.append(file_path)
                if display_sha1:
                    sha1 = sha1sum(file_path)
                    print(f"SHA1: {sha1}\t{file_path}")
                    if destination_dir is not None:
                        with open(os.path.join(destination_dir, "hashes.txt"), "a") as hashes_file:
                            hashes_file.write(f"{sha1} {file_path}\n")

    if destination_dir is not None:
        print()
        print(f"All artifacts have been written to the folder: {destination_dir}")

    return extracted_files



if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="docxTRACT",
        description=f"docxTRACT version {VERSION}",
        epilog="docxTRACT comes with ABSOLUTELY NO WARRANTY!",
    )
    parser.add_argument("filename", help="the DOCX file to extract files from")
    parser.add_argument(
        "-x",
        "--xtract",
        action="store_true",
        help="extract files (by default to current dir)",
    )
    parser.add_argument(
        "-i",
        "--img_only",
        action="store_true",
        help="extract only images",
    )
    parser.add_argument(
        "-s",
        "--sha1",
        action="store_true",
        help="display extracted file SHA1 hash",
    )
    parser.add_argument(
        "-d",
        "--destdir",
        help="define destination directory",
    )
    args = parser.parse_args()

    if not args.xtract:
        parser.print_help()
        exit()

    if not os.path.isfile(args.filename):
        print(f"File: {args.filename} does not exist.")
        exit()

    extract_docx(args.filename, args.img_only, args.sha1, args.destdir)
```

## Sample output
```
sudo python3 docxTRACT.py -x -s -d evidence sample-docx-file-for-testing.docx                                                                ─╯

░█▀▄░▄▀▀▄░█▀▄░█░█░▀▀█▀▀░▒█▀▀▄░█▀▀▄░▒█▀▀▄░▀▀█▀▀
░█░█░█░░█░█░░░▄▀▄░░▒█░░░▒█▄▄▀▒█▄▄█░▒█░░░░░▒█░░
░▀▀░░░▀▀░░▀▀▀░▀░▀░░▒█░░░▒█░▒█▒█░▒█░▒█▄▄▀░░▒█░░


 ARTIFACTS FOUND:

SHA1: dbc37923146ebf5dd595de6fb86796a9d5ed7e52	evidence/[Content_Types].xml
SHA1: 9d7abf0ee4effcecad80c8bbfb276079a05b4342	evidence/.rels
SHA1: 10da4ba054dc87be73e2beb1e899338bf97b186d	evidence/document.xml.rels
SHA1: 449b7fe314f81ceb12dab8cabdd33b64acb9fb8e	evidence/document.xml
SHA1: 230e8f4baaeb5e4a2639df9f328c847e7d3eff1d	evidence/image1.jpg
SHA1: 6d3d47e8a44dec284678d775c00bf144e67b1d3d	evidence/image3.jpeg
SHA1: f5cba8a650d6ada98d170f1b22098d93b8ff8879	evidence/theme1.xml
SHA1: 53b773f79e86ba2c91e50d6b171156f5328e1c33	evidence/image2.jpeg
SHA1: 5b7b277441f3f99d22f68c3ef099211230b64010	evidence/settings.xml
SHA1: 743020002ba05b430633bcb539b5b66f8b6674bc	evidence/fontTable.xml
SHA1: e42e4687ebd5f75dccfbecd3f9e71da80e8076d8	evidence/webSettings.xml
SHA1: 56a0d6fc7af769641aa5749d741cca80c957c96e	evidence/app.xml
SHA1: fba421f1b387a4a8499276f8cb982a0ccb481054	evidence/styles.xml
SHA1: 48a6ff63f50fafb4978d1a85e285600e663e88cb	evidence/core.xml

All artifacts have been written to the folder: evidence
```

## Disclaimer
"The scripts in this repository are intended for authorized security testing and/or educational purposes only. Unauthorized access to computer systems or networks is illegal. These scripts are provided "AS IS," without warranty of any kind. The authors of these scripts shall not be held liable for any damages arising from the use of this code. Use of these scripts for any malicious or illegal activities is strictly prohibited. The authors of these scripts assume no liability for any misuse of these scripts by third parties. By using these scripts, you agree to these terms and conditions."

## License Information

This library is released under the [Creative Commons ShareAlike 4.0 International license](https://creativecommons.org/licenses/by-sa/4.0/). You are welcome to use this library for commercial purposes. For attribution, we ask that when you begin to use our code, you email us with a link to the product being created and/or sold. We want bragging rights that we helped (in a very small part) to create your 9th world wonder. We would like the opportunity to feature your work on our homepage.
