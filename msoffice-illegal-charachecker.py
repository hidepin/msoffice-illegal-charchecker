#!/usr/bin/env python

import sys
import re
import docx

args = sys.argv

doc = docx.Document(args[1])

illegal_chars = '[０-９ａ-ｚＡ-Ｚ　！”＃＄％＆｀（）＊＋：；＜＞＝？]'

illegal_line = re.compile(r'{}'.format(illegal_chars))

for doc_paagraphs in doc.paragraphs:
    match = illegal_line.search(doc_paagraphs.text)
    if match != None:
        print(args[1] + "," + re.sub(r'({})'.format(illegal_chars), r'[\1]', doc_paagraphs.text))
