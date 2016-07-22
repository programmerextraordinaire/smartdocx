# -*- coding: latin-1 -*-
"""

docxcomments.py

--tkp

Extract comments from .docx MS Word file(s)

"""


import argparse
import glob
import os
import sys

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
ID = WORD_NAMESPACE + 'id'
DATE = WORD_NAMESPACE + 'date'
AUTHOR = WORD_NAMESPACE + 'author'
COMMENT = WORD_NAMESPACE + 'comment'


def get_docx_comments(path, anonymous):
    """
    Take the path of a docx file as argument, return the comments text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read(r'word/comments.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for comment in tree.getiterator(COMMENT):
        id = int(comment.attrib[ID]) + 1    # Bump ID number cuz people don't like zero.
        dt = comment.attrib[DATE]
        author = comment.attrib[AUTHOR]

        for paragraph in comment.getiterator(PARA):
            texts = [node.text
                     for node in paragraph.getiterator(TEXT)
                     if node.text]
            if texts:
                paragraphs.append('{0}. Author: {1}  Date:{2}'.format(id, author, dt))
                paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)


def main(path, anonymous, outfname):
    searchpath = os.path.join(path, "*.docx")
    files = glob.glob(searchpath)
    for filename in files:
        print(filename)
        comments = get_docx_comments(filename, anonymous)
        print(comments)
        


# Example usage: python docxcomments.py --doctype protocol 2
if __name__ == "__main__":
    # Get arguments
    parser = argparse.ArgumentParser(description='Parse comments from MS Word documents.')
    parser.add_argument('--anonymous', action='store_true', help='make comments anonymous?')
    parser.add_argument('--path', default=os.getcwd(), help='input path for documents')
    parser.add_argument('--filename', default=None, help='export comments to this file.')
    args = parser.parse_args()
    #args = parser.parse_args(['--a', '--path', r'C:\Python\Scripts\AppVizo\smartdocx', '--f', ''])
    main(args.path, args.anonymous, args.filename)
    
    