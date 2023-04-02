#!/usr/bin/env python3

import argparse
import hashlib
import os
import zipfile

# Print ASCII art
print("""
░█▀▄░▄▀▀▄░█▀▄░█░█░▀▀█▀▀░▒█▀▀▄░█▀▀▄░▒█▀▀▄░▀▀█▀▀
░█░█░█░░█░█░░░▄▀▄░░▒█░░░▒█▄▄▀▒█▄▄█░▒█░░░░░▒█░░
░▀▀░░░▀▀░░▀▀▀░▀░▀░░▒█░░░▒█░▒█▒█░▒█░▒█▄▄▀░░▒█░░
""")

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
