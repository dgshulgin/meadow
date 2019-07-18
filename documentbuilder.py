# -*- encoding: utf-8 -*-

import pydoc
import os.path
import logging

'''
Диспетчер плагинов для создания документов определенного типа
по шаблону. 
Плагин может быть
зарегистрирован для обработки нескольких типов расширений.
Диспетчер анализирует список расширений, переданных при запуске сервера, 
и формирует список плагинов.
Плагин должен быть размещен в файле с названием <ext>builder.py.
Плагин должен наследовать класс DocumentBuilder и реализовать метод build
с двумя параметрами:
- template - наименование документа-шаблона
- xml_data - данные в формате XML для заполнения документа-шаблона

'''

class DocumentBuilder(object):

    def __init__(self, config):
        self._handlers = {}
        #
        self.__logger = logging.getLogger('mofserver')
        #
        self._folder_templ    = config["folder_templ"]
        self._folder_out      = config["folder_out"]
        path_p = os.path.join((os.curdir, os.sep))
        self._folder_handlers = path_p #os.path.normpath(path_p)
        #
        self.__load_plugins(config["formats"])

    '''
    '''    
    def __load_plugins(self, plugins_list):
        for pln in plugins_list:
            mod_name = "{0}{1}builder.{2}Builder".format(self._folder_handlers, \
                                                        pln.lower(), \
                                                        pln.upper())
            klasse = pydoc.locate(mod_name)
            if klasse is not None:
                self.register(pln, klasse)
                #self._handlers[pln] = klasse #Перед использованием надо создать экземпляр klasse

    '''
    TODO 
    '''
    def register(self, ext, plg_klasse):
        self._handlers[ext.upper()] = plg_klasse

    def getPluginClass(self, template):
        '''
        Возвращает экземпляр класса-плагина с обязательным методом run, либо None.
        '''
        pklasse = None
        ext = os.path.splitext(template)[1][1:].upper()
        if ext in self._handlers.keys():
            pklasse = self._handlers[ext]
        return pklasse

    def build(self, template, out_name, xml_file):
        #

        #Формируем полные пути для документов
        path_t = os.path.join((os.curdir, self._folder_templ, template))
        path_t = os.path.normpath(path_t)

        path_o = os.path.join((os.curdir, self._folder_out, out_name))
        path_o = os.path.normpath(path_o)

        #c = self.getPluginClass(path_t)  #get plugin by template extension
        c = self.getPluginClass(path_o)   #get plugin by output document extension 
        if c is not None:
            c1 = c()
            return c1.build(path_t, path_o, xml_file)
        else:
            #msg = str('There is no handler for {0}'.format(path_t))
            msg = str('There is no handler for {0}'.format(path_o))
            self.__logger.warning(msg)
            return msg
        #return self._handlers[template].build(template, xmldata)