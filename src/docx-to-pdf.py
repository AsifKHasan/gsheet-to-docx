#!/usr/bin/env python3
'''
usage:
python docx-to-pdf.py
'''
import os
import win32com.client as client

def convert_to_pdf(doc):
    doc_path = os.path.abspath(doc)
    try:
        word = client.DispatchEx("Word.Application")
        pdf_path = doc_path.replace(".docx", r".pdf")
        worddoc = word.Documents.Open(doc_path)
        for toc in worddoc.TablesOfContents:
            toc.Update()

        worddoc.Save()
        worddoc.SaveAs(pdf_path, FileFormat = 17)
        worddoc.Close()
    except Exception as e:
            raise e
    finally:
            word.Quit()

convert_to_pdf("../out/DSHE-DG94__8__sow-specifications-compliance.docx")
