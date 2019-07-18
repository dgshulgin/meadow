# -*- encoding: utf-8 -*-

from documentbuilder import DocumentBuilder

class PDFBuilder(DocumentBuilder):
    def __init__(self):
        pass
    def build(self, template, out_name, xmldata):
        return str('Building PDF')