#!/usr/bin/env python

import sys
import os
import re
import docx

illegal_chars = '[０-９ａ-ｚＡ-Ｚ　！”＃＄％＆｀（）＊＋：；＜＞＝？]'

def check_docx(filename):
    docx = docx.Document(filename)
    illegal_line = re.compile(r'{}'.format(illegal_chars))
    for doc_paagraphs in docx.paragraphs:
        match = illegal_line.search(doc_paagraphs.text)
        if match != None:
            print(args[1] + "," + re.sub(r'({})'.format(illegal_chars), r'[\1]', doc_paagraphs.text))

if __name__ == '__main__':
    args = sys.argv
    base, ext = os.path.splitext(args[1])

    if ext == '.docx' or ext == '.doc':
        check_docx(args[1])
