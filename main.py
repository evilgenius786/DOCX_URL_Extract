import csv
from datetime import datetime

import win32com.client
import os
from docx2python import docx2python

from urlextract import URLExtract

indir = "files"
outfile = "out.csv"
headers = ["File", "URL"]
word = win32com.client.Dispatch("Word.Application")
word.visible = 0


def main():
    for docx in os.listdir(indir):
        if docx.lower().endswith(".docx") and not docx.startswith("~"):
            getFooter(docx)


def getFooter(file):
    print(datetime.now(), "Extracting URLs from", file)
    if not os.path.isfile(outfile):
        with open(outfile, 'w', encoding='utf8', newline='') as ofile:
            csv.DictWriter(ofile, fieldnames=headers).writeheader()
    extractor = URLExtract()
    footnotes = docx2python(indir + "/" + file).footnotes_runs
    print(file, footnotes)
    urls = extractor.find_urls(str(docx2python(indir + "/" + file).footnotes_runs))
    print(file, urls)
    with open(outfile, 'a', encoding='utf8', newline='') as ofile:
        f = csv.DictWriter(ofile, fieldnames=headers)
        for url in urls:
            f.writerow({
                "File": file,
                "URL": url
            })


def convert(d, pdf):
    f = os.path.abspath(f"{d}/{pdf}")
    wb = word.Documents.Open(f)
    wb.SaveAs2(f.replace(".pdf", ".docx"), FileFormat=16)
    print(datetime.now(), "Saved as", f.replace(".pdf", ".docx"))
    wb.Close()
    word.Quit()


def pdftoword():
    print(datetime.now(), "Converting all files from PDF to DOCX first.")
    for pdf in os.listdir(indir):
        if pdf.lower().endswith(".pdf"):
            print(datetime.now(), "Working on", pdf)
            print(datetime.now(), indir, pdf)
            return


def logo():
    os.system("color 0a")
    print(r"""
    _________ .__                .___   ________                     
    \_   ___ \|  |__ _____     __| _/  /  _____/_____ _______ ___.__.
    /    \  \/|  |  \\__  \   / __ |  /   \  ___\__  \\_  __ <   |  |
    \     \___|   Y  \/ __ \_/ /_/ |  \    \_\  \/ __ \|  | \/\___  |
     \______  /___|  (____  /\____ |   \______  (____  /__|   / ____|
            \/     \/     \/      \/          \/     \/       \/     
==========================================================================
            Footnotes URL extractor from PDF and DOCX by:
                http://github.com/evilgenius786
==========================================================================
[+] Process DOCX
[+] Convert PDF to DOCX
[+] Extract only footnote URLS
[+] Work with hyperlinks
__________________________________________________________________________
""")


if __name__ == '__main__':
    logo()
    # main()
    choice = input(f"Enter 1 to converto PDF to DOCX, 2 to extract URLs from DOCX (from dir ./{indir}): ")
    if choice == "1":
        convert("./", "pdf.pdf")
    else:
        main()
