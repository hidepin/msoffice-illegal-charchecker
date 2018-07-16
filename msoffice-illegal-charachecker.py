#!/usr/bin/env python

import docx
import re

doc = docx.Document('sample.docx')

illegal_line = re.compile(r'[０-９ａ-ｚＡ-Ｚ]')

for doc_paagraphs in doc.paragraphs:
    match = illegal_line.search(doc_paagraphs.text)
    if match != None:
        print(doc_paagraphs.text)
