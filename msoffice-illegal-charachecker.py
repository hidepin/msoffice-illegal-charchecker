#!/usr/bin/env python

import sys
import re
import docx

args = sys.argv

doc = docx.Document(args[1])

illegal_line = re.compile(r'[０-９ａ-ｚＡ-Ｚ]')

for doc_paagraphs in doc.paragraphs:
    match = illegal_line.search(doc_paagraphs.text)
    if match != None:
        print(args[1] + "," + doc_paagraphs.text)
