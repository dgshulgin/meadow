#! /usr/bin/python
# -*- encoding: utf-8 -*-

'''
Сервер генерации документов XOD/OOXML/ODF/PDF на основе документа-шаблона.
Для заполнения шаблона используются данные в формате XML.
При запуске сервера должны быть переданы обязательные параметры:
- хост -h<host-name>
- порт -p<port>
- расширения создаваемых документов -f<ext>
Может быть переданно несколько параметров -f, например
./s2.py -h myoffice.ru -p 8000 -f XODT -f PDF -f XODS -f DOCX

Необязательные параметры:
- путь к каталогу шаблонов
- путь к каталогу выходных документов

Вызовы:
make - формирование документа по шаблону и набору данных
listMethods - предотавляет список доступных вызовов (методов) RPC-сервера
methodHelp - справка по указанному методу

Components:
s2.py              - RPC server starter, this file
rpcdispatcher.py   - RPC server commands dispatching
documentbuilder.py - Plugins dispatcher. Plugin is responsible to build
                     a document of specified format XODT, XODS, PDF, etc.
xodbuilder.py      - Plugin builds an XODT document


Руководствуясь параметрами командной строки сервер запустит необходимые плагины. 

1. Должен быть конфигурационный файл в котором каждому шаблону ставится в соотв плагин-обработчик
2. В этом же конфигурационном файлк хранятся сетевые настройки 
3. Плагин выбирается и запускается соотв имени шаблона

'''


from SimpleXMLRPCServer import SimpleXMLRPCServer, list_public_methods
from SimpleXMLRPCServer import SimpleXMLRPCRequestHandler
import logging
import argparse

from rpcdispatcher import RPCDispatcher


def main():

    ## Cmd line parameters
    parser = argparse.ArgumentParser(description='MyOffice Document API. Сервер генерации документов.')
    parser.add_argument('-p', action="store", dest="port", \
                        default='8000', type=int, \
                        help='Порт. Значение по умолчанию 8000')
    parser.add_argument('-s', action="store", dest="host", default='127.0.0.1', \
                        help='Хост. Значение по умолчанию 127.0.0.1')
    parser.add_argument('-t', action="store", dest="folder_templ", default='./template', \
                        help='Путь к каталогу шаблонов. Значение по умолчанию ./template')
    parser.add_argument('-o', action="store", dest="folder_out", default='./out', \
                        help='Путь к каталогу целевых документов. Значение по умоланию ./out')                
    #parser.add_argument('-d', '--debug', action='store_true',        help='Enable debug mode with debug output')
    parser.add_argument('-f', action="append", dest="formats", default=["XODT", "PDF"], \
                        help='Форматы документов-шаблонов. Используйте несколько ключей чтобы перечислить все необходимые форматы',)

    args = parser.parse_args()

    logger = logging.getLogger('mofserver')

    server = SimpleXMLRPCServer((args.host, args.port), allow_none = True, logRequests = True)
    server.register_introspection_functions() # для удаленных вызовов listMethods, methodHelp
    #
    config = {}
    config["folder_templ"] = args.folder_templ
    config["folder_out"]   = args.folder_out
    config["formats"]      = []
    #config["log"]          = logging

    for p in args.formats:
        config["formats"].append(p.upper())

    if len(config["formats"]) < 1:
        logger.warning('Форматы для обработки не указаны.')

    server.register_instance(RPCDispatcher(config))
    #
    try:
        logger.info('Для завершения работы сервера нажмите Ctrl-C')
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Сервер завершает работу...")
    
    logging.shutdown()
    server.shutdown()

if __name__ == "__main__" : main()