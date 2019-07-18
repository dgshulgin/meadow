# -*- encoding: utf-8 -*-

'''
Диспетчер команд RPC-сервера.

Вызовы:
make - формирование документа по шаблону и набору данных
listMethods - предотавляет список доступных вызовов (методов) RPC-сервера
methodHelp - справка по указанному методу

'''

import inspect

from SimpleXMLRPCServer import SimpleXMLRPCServer, list_public_methods

#from documentbuilder  import DocumentBuilder
import documentbuilder
import logging

class RPCDispatcher(object):

    def __init__(self, config):
        self.__master = documentbuilder.DocumentBuilder(config)
        self.__logger = logging.getLogger('mofserver') #config["log"]

    def _listMethods(self):
        return list_public_methods(self)

    def _methodHelp(self, method):
        f = getattr(self, method)
        return inspect.getdoc(f)

    '''
    Функция удаленного вызова make.
    Формирование итогового документа по шаблону template используя данные xmldata.
    Параметры:
    template - наименование документа-шаблона. Ищет документ в каталоге 
    out_name - наименование выходного документа, без расширения. Расширение приделает 
    плагин, ответственный за формирование выходного документа.
    xml_file - файл данных в формате XML, тегами являются наименования соотв закладок
    в документе-шаблоне. 
    '''
    def make(self, template, out_name, xml_file):
        self.__logger.info('Удаленный вызов make')
        '''
        Построить документ по указанному шаблону template, используя данные xmldata.
        '''
        return self.__master.build(template, out_name, xml_file)