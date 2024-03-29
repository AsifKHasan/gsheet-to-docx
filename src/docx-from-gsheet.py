#!/usr/bin/env python3
'''
usage:
./docx-from-gsheet.py --config '../conf/config.yml'
python docx-from-gsheet.py --config "../conf/config.yml"
pip list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1  | xargs -n1 pip install -U

[soffice/libreoffice] --headless --convert-to pdf docx-to-convert.docx
'''
import os
import sys
import json
import importlib
import time
import yaml
import datetime
import argparse
import pprint
if sys.platform == 'win32':
	import win32com.client as client

from docx import Document

from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.docx.docx_helper import DocxHelper
from helper.docx.docx_util import *

class DocxFromGsheet(object):

	def __init__(self, config_path, gsheet=None):
		self.start_time = int(round(time.time() * 1000))
		self._config_path = os.path.abspath(config_path)
		self._data = {}
		self._gsheet = gsheet

	def update_toc(self, docx_path, generate_pdf):
		doc_path = os.path.abspath(docx_path)
		try:
			word = client.DispatchEx("Word.Application")
			worddoc = word.Documents.Open(doc_path)
			# for toc in worddoc.TablesOfContents:
			# 	toc.Update()
			#
			# worddoc.Save()

			if generate_pdf:
				pdf_path = doc_path.replace(".docx", r".pdf")
				worddoc.SaveAs(pdf_path, FileFormat = 17)

			worddoc.Close()
		except Exception as e:
			raise e
		finally:
			word.Quit()

	def generate_docx(self):
		for section in self._data['sections']:
			content_type = section['content-type']

			# force table formatter for gsheet content
			if content_type == 'gsheet': content_type = 'table'

			module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
			module.generate(section, self._doc, self._docxhelper._sections, self._CONFIG)

		self._doc.save(self._CONFIG['files']['output-docx'])
		set_updatefields_true(self._CONFIG['files']['output-docx'])

		if sys.platform == 'win32' and self._CONFIG['docx-related']['update-toc']:
			self.update_toc(self._CONFIG['files']['output-docx'], self._CONFIG['docx-related']['generate-pdf'])

	def run(self):
		self.set_up()
		# process gsheets one by one
		for gsheet in self._CONFIG['gsheets']:
			self._data = self._gsheethelper.process_gsheet(gsheet)

			self._CONFIG['files']['output-json'] = os.path.abspath('{0}/{1}.json'.format(self._CONFIG['dirs']['output-dir'], gsheet))
			self.save_json()

			# docx-helper
			self._CONFIG['files']['output-docx'] = os.path.abspath('{0}/{1}.docx'.format(self._CONFIG['dirs']['output-dir'], gsheet))
			self._docxhelper = DocxHelper(self._CONFIG['files']['docx-template'], self._CONFIG['files']['docx-styles'], self._CONFIG['files']['output-docx'])
			self._doc = self._docxhelper.init()
			self.generate_docx()

			self.tear_down()

	def set_up(self):
		# configuration
		self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
		config_dir = os.path.dirname(self._config_path)

		# if gsheet name was provided as parameter, override the configuration
		if self._gsheet:
			self._CONFIG['gsheets'] = [self._gsheet]

		self._CONFIG['dirs']['data-dir'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['dirs']['data-dir']))
		self._CONFIG['dirs']['output-dir'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['dirs']['output-dir']))
		self._CONFIG['dirs']['temp-dir'] = os.path.abspath('{0}/tmp'.format(self._CONFIG['dirs']['output-dir']))
		if not os.path.exists(self._CONFIG['dirs']['temp-dir']):
			os.makedirs(self._CONFIG['dirs']['temp-dir'])

		self._CONFIG['files']['google-cred'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['files']['google-cred']))
		self._CONFIG['files']['docx-styles'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['files']['docx-styles']))
		self._CONFIG['files']['docx-template'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['files']['docx-template']))

		# gsheet-helper
		self._gsheethelper = GsheetHelper()
		self._gsheethelper.init(self._CONFIG)

	def save_json(self):
		with open(self._CONFIG['files']['output-json'], "w") as f:
			f.write(json.dumps(self._data, sort_keys=False, indent=4))

	def tear_down(self):
		self.end_time = int(round(time.time() * 1000))
		debug("Script took {} seconds".format((self.end_time - self.start_time)/1000))

if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-g", "--gsheet", required=False, help="gsheet name to override gsheet list provided in configuration")
	args = vars(ap.parse_args())

	generator = DocxFromGsheet(args["config"], args["gsheet"])
	generator.run()
