#!/usr/bin/env python3
'''
from command line
------------------
./latex2pdf.py --latexfile ./salary-advice.tex --datafile ./out/salary-advice_Oct18.json --pdffile ./out/salary-advice_Oct18.pdf
./latex2pdf.py --latexfile ./salary-summary-for-management.tex --datafile ./out/salary-summary-for-management_Oct18.json --pdffile ./out/salary-summary-for-management_Oct18.pdf

from py files
------------------
pdfgenerator = Latex2Pdf(latexfile, pdffile)
pdfgenerator.generate_pdf(data_as_json)
'''
import os
import shutil
import sys
import json
import time
import argparse

from jinja2 import Environment, FileSystemLoader
from latex.jinja2 import make_env
from latex.build import LatexMkBuilder

def format_currency(amount):
    if not amount:
        return "0.00"
    return "{:,.2f}".format(round(amount, 2))

class Latex2Pdf(object):

    _ENV = Environment(
        block_start_string='{%',
        block_end_string='%}',
        variable_start_string='{{%',
        variable_end_string='%}}',
        loader=FileSystemLoader(".")
        )

    _ENV.filters['format_currency'] = format_currency


    def __init__(self, latexfile, pdffile):
        self.start_time = int(round(time.time() * 1000))
        self._TEMPLATE = latexfile
        self._OUTPUT_PDF = pdffile

    def generate_pdf(self, data):
        self._data = data
        template = self._ENV.get_template(self._TEMPLATE)
        builder = LatexMkBuilder(variant='xelatex')

        pdf = builder.build_pdf(template.render(data=self._data))
        pdf.save_to(self._OUTPUT_PDF)
        print("Pdf successfully gerenated : ", self._OUTPUT_PDF)

    def set_up(self, datafile):
        with open(datafile) as df:
            self._data = json.load(df)

    def tear_down(self):
        self.end_time = int(round(time.time() * 1000))
        print("Script took {} seconds".format((self.end_time - self.start_time)/1000))

    def run(self, datafile):
        self.set_up(datafile)
        self.generate_pdf(self._data)
        self.tear_down()

if __name__ == '__main__':
    # construct the argument parse and parse the arguments
    ap = argparse.ArgumentParser()
    ap.add_argument("-l", "--latexfile", required=True, help="Latex template to generate pdf output")
    ap.add_argument("-d", "--datafile", required=True, help="Input datafile in json format")
    ap.add_argument("-p", "--pdffile", required=True, help="Output pdf file to be generated")
    args = vars(ap.parse_args())

    generator = Latex2Pdf(args["latexfile"], args["pdffile"])
    generator.run(args["datafile"])
