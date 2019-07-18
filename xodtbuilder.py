# -*- encoding: utf-8 -*-

'''
Плагин для создания документов XODT.

Алгоритм воспроизведен по коду ЭОС.
Шаблон документа содержит таблицу подписантов. Выходной документ содержит:
таблицу подписантов,


1. отключить отслеживание изменений т.к. будем модифицировать документ, //TracksOnOff(true); 
2. Сформировать таблицу подписантов.  //CreateSignerTable();
2.1 Варианты: а) Шаблон содержит закзадку LISTSIGNERSSTAMPS б) Шаблон сореджит готовую 
таблицу подписантов в закладками в полях в)Шаблон содержит закзадку LISTSIGNERSSTAMPS 
и таблицу подписантов со своими закладками.


2.2. Считаем кол-во место подписи (они уже есть в таблице подписантов).
2.3. Если кол-во мест подписи (пока пустых) соотв или превышает кол-во переданных данных, то вставлять таблицу вместо LISTSIGNERSSTAMPS ненужно. 
2.4. Вставляем таблицу в букмарк LISTSIGNERSSTAMPS, Форматируем ширину таблицы и удаляем закладку
2.5. Вставляем в таблицу значени полей - это разниа между имеющимися местами подписи и переданными данными.



3. Сформировать таблицу адресатов. //CreateAddrTable();
3.1 шаблон содержит закладку LISTADDR, внутри которой надо сформировать таблицу адресатов
3.2 закладка LISTADDR отсуствует и в таким случае считаем, что таблица адресатов формируется 
другим способом - сама таблица присутсвутет в тексте и в ячейках есть закладки, которые надо заполнять


4. Заполнить закладки в таблицах подписантов и адресатов, а также ряд общих закладок данными, 
переданными в структуре XML. //FillBookmarks();
4.1 Удалить закладки, сохранив значимые данные. 
5. Удалить в таблицах подписантов и адресатов незаполненные закладки. Таковые образуются потому, что данных было
переданно меньше, чем знакомест в шаблоне. //RemoveUnusedBookmarks();

6. Оставшиеся в документе закладки заменить на текстовые метки формата [Bookmark name], 
имя закладки в верхнем регистре.  //SetMarks();
7.  Вставляет угловой штамп. Если в документе есть закладка STAMPCORNER, то вставляем 
штамп (ищображение). //StampCorner(Path.GetDirectoryName(Xml));

8. В  структуре данных  могут быть переданы данные в фомрмате ИмяЗакладки-ЗначениеЗакладки для полей
которые заранее определить невозможно. Нужно пройтись по документу и заменить такие закладки в документе
(если они есть) на значимые данные. //SetOtherMarks();
 
9. Включить отслеживание изменений, формирование документа завершено.  //TracksOnOff(false); //включить отслеживание изменений
'''

import abc
import threading
import os.path
import logging
import declxml as xml
import re 
import documentbuilder
from MyOfficeSDK import CoreAPI as mof

class Executor(object):
    def __init__(self):
        self.name = None
        self.phone = None

class Annotation(object):
    def __init__(self):
        self.text = None

class Signer(object):
    def __init__(self):
        self.due = None
        self.post = None
        self.name = None

class AbstractEntryBuilder:
    __metaclass__ = abc.ABCMeta
    
    def replace_bookmark(self, doc, tag, value):
        '''
        Заменить содержимое закладки с именем tag на значение value, затем удалить
        закладку tag. Содержимое закладки остается в документе.
        '''
        bmk_range = doc.getBookmarks().getBookmarkRange(tag)
        if bmk_range is not None:
            bmk_range.replaceText(value.encode('utf8'))
            doc.getBookmarks().removeBookmark(tag)
        else:
            #Закладки с таким имененем в документе нет, попробуем найти
            #  текстовую метку
            search = mof.createSearch(doc)
            rngs = search.findText(tag)
            for r in rngs:
                r.replaceText(value.encode('utf8'))

    
    @abc.abstractmethod
    def update(self):
        '''
        Обновить значение закладки с предопределенным именем на значение, полученное
        из данных XML.
        '''        
        pass

class DummyEntryBuilder(AbstractEntryBuilder):
    def __init__(self, doc, tag, value):
        self.doc   = doc
        self.tag   = tag
        self.value = value
    
    def update(self):
        super(DummyEntryBuilder, self).replace_bookmark(self.doc, self.tag, self.value)

class ExecutorEntryBuilder(AbstractEntryBuilder):
    def __init__(self, doc, xml_data):
        self.doc      =  doc
        self.xml_data = xml_data
 
    def update(self):
        logging.debug('ExecutorEntryBuilder')
        executor_proc = xml.user_object('PASSPORT/EXECUTOR', Executor, [
            xml.string('EXECUTORNAME',  alias='name'),
            xml.string('EXECUTORPHONE', alias='phone')
        ])
        executor = xml.parse_from_file(executor_proc, self.xml_data)
        super(ExecutorEntryBuilder, self).replace_bookmark(self.doc, 'EXECUTORNAME',  executor.name)
        super(ExecutorEntryBuilder, self).replace_bookmark(self.doc, 'EXECUTORPHONE', executor.phone)
 
class AnnotationEntryBuilder(AbstractEntryBuilder):
    def __init__(self, doc, xml_data):
        self.doc =  doc
        self.xml_data = xml_data

    def update(self):
        logging.debug('AnnotationEntryBuilder')
        annotation_proc = xml.user_object('PASSPORT', Annotation, [
            xml.string('ANNOTATION', alias='text')
        ])
        annotation = xml.parse_from_file(annotation_proc, self.xml_data)
        super(AnnotationEntryBuilder, self).replace_bookmark(self.doc, 'ANNOTATION', annotation.text)

class AddrTableBuilder(AbstractEntryBuilder):
    def __init__(self, doc, xml_data):
        self.doc      = doc
        self.xml_data = xml_data   

    def update(self):
        pass

class SignersTableBuilder(AbstractEntryBuilder):
    def __init__(self, doc, xml_data):
        self.doc      = doc
        self.xml_data = xml_data

    def __get_fixed_placeholders(self, names_wanted):
        '''
Построить список подтвержденных имен букмарков для вставки данных 
подписантов, а также кол-во подписантов, уже обеспеченных букмарками
в шаблоне.
Параметры:
app - Объект приложения MyOffice CoreAPI, для доступа к структуре
        документа-шаблона.
names_wanted - список шаблонов-наименований для букмарков,
                типа SIGNERNAME, SIGNERPOST
Возвращает:
max_index - кол-во подписантов для которых в шаблоне уже имеются 
            букмарки
exist - список подтвержденных имен букмарков в шаблоне, для вставки 
        данных подписантов
        '''
        bookmarks = self.doc.getBookmarks()
        '''
        TODO Сейчас нет возможности получить точное количество букмарков в
        документе, поэтому мы предполагаем что подписантов м.б. не более 30 и ищем
        букмарки с именами типа SIGNERDUE1-SIGNERDUE30
        '''
        # Формируем список всех возможных имен букмарков
        wanted = ["{0}{1}".format(n, bi) for bi in range(1, 31) for n in names_wanted]
        exist = []
        for bmk_name in wanted:
            # Подтверждаем наличие букмарка с таким именем в документе-шаблоне
            bmk_rng = bookmarks.getBookmarkRange(bmk_name)
            if bmk_rng is not None:
                exist.append(bmk_name)
        for txt_name in wanted:
            search = mof.createSearch(self.doc)
            rngs = search.findText("[{0}]".format(txt_name))
            if rngs is not None:
                for r in rngs:
                    exist.append("[{0}]".format(txt_name))
        '''
        exist содержит список подтвержденных имен, возьмем последний элемент
        и вырежем из него номер.
        '''
        max_index = 0
        if len(exist):
            #max_index = int(re.sub('[a-zA-Z]', '', exist[-1]))
            max_index = int(filter(str.isdigit,exist[-1]))
        return max_index, exist

    def __get_data_items(self):
        '''
        Построить список элементов данных, для вставки в указанные позиции 
        (закладки).
        Параметры:
        нет
        Возвращает:
        max_index - кол-во элементов данных
        exist - список элементов данных
        '''
        signers_proc = xml.dictionary('PASSPORT/SIGNERLIST', [
        xml.array(xml.user_object('SignerInfo', Signer, [
                xml.string('SIGNERDUE',  alias='due'),
                xml.string('SIGNERPOST', alias='post'),
                xml.string('SIGNERNAME', alias='name')
            ], alias='slist'))
        ])
        s = xml.parse_from_file(signers_proc,  self.xml_data)
        return s['slist']

    def __insert_table(self, range_, num_rows, num_cols):
        '''
        TODO настраиваются невидимые границы и ширины столбцов
        TODO букмарк-плейсхолдер удаляется
        '''
        #
        t_name = range_.extractText()
        #Есть нюанс: insertTable по умолчанию добавляет порядковый номер к имени таблицы
        #при вставке. 
        t_id = range_.getBegin().insertTable(num_rows, num_cols, t_name)
        t_obj = self.doc.getBlocks().getTable(t_id)
        return t_obj

    def update(self):
        '''
2. Сформировать таблицу подписантов.  //CreateSignerTable();
2.1 Варианты: а) Шаблон содержит закзадку LISTSIGNERSSTAMPS б) Шаблон сореджит готовую 
    таблицу подписантов в закладками в полях в)Шаблон содержит закзадку LISTSIGNERSSTAMPS 
    и таблицу подписантов со своими закладками.

2.2. Считаем кол-во место подписи (они уже есть в таблице подписантов).
2.3. Если кол-во мест подписи (пока пустых) соотв или превышает кол-во переданных данных, 
    то вставлять таблицу вместо LISTSIGNERSSTAMPS ненужно. 
2.4. Вставляем таблицу в букмарк LISTSIGNERSSTAMPS, Форматируем ширину таблицы и удаляем закладку
2.5. Вставляем в таблицу значени полей - это разниа между имеющимися местами подписи и переданными 
    данными.
        '''
        '''
        placeholders_list
        signers_list
        if(len(signers) > len(placeholders)):
            if(check_bookmark("LISTSIGNERSSTAMPS")):
                insert_table
                insert_table_bookmarks(signers - placeholders), update_placeholders()
                delete_bookmark
        for s, p in signers, placeholders:
            fill_entry(p, s)
        '''        
        logging.debug('SignersTableBuilder')
        #Список всех шаблонов для вставки инфо о подписантах.
        fixed_plh_count, fixed_plh_list = \
                self.__get_fixed_placeholders(['SIGNERPOST', \
                                            'SIGNERNAME', \
                                            'SIGNERSTAMP'])
        #Сейчас fixed_plh_list содержит SIGNERDUE1..N,SIGNERPOST1..N, 
        # SIGNERNAME1..N соотв имеющимся именам закладок
        #
        signers_list = self.__get_data_items()
        signers_count = len(signers_list)
        #
        bookmarks = self.doc.getBookmarks()
        if(signers_count > fixed_plh_count):
            #Подписантов больше чем сидячих мест, вставим доп таблицу...
            signers_plh_name  = 'LISTSIGNERSSTAMPS'
            signers_plh_table = bookmarks.getBookmarkRange(signers_plh_name)
            if signers_plh_table is not None:
                #...если есть куда.
                #Вставляем таблицу для дополнительных подписантов. Для каждого подписанта
                #вставляется блок 2х2 ячейки, ячейки во второй строке объединены.
                t_signers = self.__insert_table(signers_plh_table, (signers_count - fixed_plh_count)*2, 2)
                #
                t_signers.setColumnWidth(0, 325) #20pt
                t_signers.setColumnWidth(1, 122) #10pt
                #
                #for r in range(0, (signers_count-fixed_plh_count)*2, 2):
                for r in range(0, signers_count-fixed_plh_count):
                    #должность
                    cell_post = t_signers.getCell(mof.CellPosition(r,0))
                    #cell_post.getRange().getBegin().insertBookmark('SIGNERPOST{0}'.format(str(fixed_plh_count + r + 1)))
                    #workaround
                    cell_post.getRange().getBegin().insertText("[SIGNERPOST{0}]".format(str(fixed_plh_count + r + 1)))
                    #ФИО
                    cell_fio = t_signers.getCell(mof.CellPosition(r,1))
                    #cell_fio.getRange().getBegin().insertBookmark('SIGNERNAME{0}'.format(str(fixed_plh_count + r + 1)))
                    #workaround
                    cell_fio.getRange().getBegin().insertText("[SIGNERNAME{0}]".format(str(fixed_plh_count + r + 1)))
                    #Штамп
                    cr = t_signers.getCellRange(mof.CellRangePosition(r+1,0,r+1,1))
                    cr.merge()
                    cell_stamp = t_signers.getCell(mof.CellPosition(r+1,0))
                    #cell_stamp.getRange().getBegin().insertBookmark('SIGNERSTAMP{0}'.format(str(fixed_plh_count + r + 1)))
                    #workaround
                    cell_stamp.getRange().getBegin().insertText("[SIGNERSTAMP{0}]".format(str(fixed_plh_count + r + 1)))
                #
                signers_plh_table.replaceText('')
                self.doc.getBookmarks().removeBookmark(signers_plh_name)
            else:
                logging.warning('XOD Worker, not enough room to insert all signers. Bookmark LISTSIGNERSSTAMPS is absent')
        #Теперь заменим все закладки подписантов в списке fixed_plh_count на значимые 
        #данные из списка signers_count, и удалим все закладки подписантов.
        #
        #Обновление списка шаблонов для вставки инфо о подписантах. 
        fixed_plh_count, fixed_plh_list = \
                self.__get_fixed_placeholders(['SIGNERPOST', \
                                            'SIGNERNAME', \
                                            'SIGNERSTAMP'])        
        for plh in fixed_plh_list:
            #data_index = int(re.sub('[a-zA-Z]', '', plh))
            data_index = int(filter(str.isdigit,plh))
            if 'SIGNERPOST' in plh:
                fill_entry(DummyEntryBuilder(self.doc, plh, signers_list[data_index-1].post))
            elif 'SIGNERNAME' in plh:
                fill_entry(DummyEntryBuilder(self.doc, plh, signers_list[data_index-1].name))
        
          
def fill_entry(builder):
    builder.update()

def worker(path_t, path_o, xml_file):
    logging.debug("XOD Worker for template %s started", path_t)
    #
    app      = mof.Application()
    app.doc  = app.loadDocument(path_t)
    
    fill_entry(SignersTableBuilder(app.doc, xml_file))
    '''
    app.doc = fill_table(AddrTableBuilder(app.doc, xml_file))
    '''

    fill_entry(ExecutorEntryBuilder(app.doc, xml_file))
    fill_entry(AnnotationEntryBuilder(app.doc, xml_file))

    '''
    app.doc = fill_entry(RefNumDateEntryBuilder(app.doc, xml_file))
    app.doc = fill_entry(RefCrpNumDateEntryBuilder(app.doc, xml_file))
    app.doc = fill_entry(AccostEntryBuilder(app.doc, xml_file))

    app.doc = fill_entry(CornerStampEntryBuilder(app.doc, xml_file))

    app.doc = fill_list(OtherEntriesBuilder(app.doc, xml_file))

    app.doc = clear_unused_bookmarks(app.doc)
    '''
    #
    app.doc.saveAs(path_o)
    #
    logging.debug("XOD Worker for template %s done with %s.", path_t, path_o)
#

class XODTBuilder(documentbuilder.DocumentBuilder):
    def __init__(self):
        self.__plugin_ext = "xodt"
       
    '''
    Входные значения:
    path_temp - путь к файлу шаблона
    path_out - путь в результирующему документу
    xml_data - данные для заполнения шаблона в формате XML
    '''
    def build(self, path_temp, path_out, xml_data):
        logging.basicConfig(level=logging.DEBUG, format='%(relativeCreated)6d %(threadName)s %(message)s')
        t = threading.Thread(target=worker, args=(path_temp, path_out, xml_data))
        t.start()
        return str('OK')