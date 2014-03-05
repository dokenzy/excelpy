# -*- coding: utf-8 -*-
import os
import shutil
from lxml import etree
from zipfile import ZipFile
import fnmatch
import json
import re

UTF_8 = 'utf-8'

from .xlsx_ns import NS_SPREADSHEETML, NS_CONTENT_TYPES, NS_PROPERTIES, NS_DOC_PROPS_VTYPES,\
    NS_RELS, NS_REL_WORKSHEET, NS_WORKSHEET_R


class ExcelPy(object):
    ''' ExcelPy class
    '''
    def __init__(self, excel_file_path):
        self.last_worksheet_num = 0  # this is required when editing [Content_Type].xml
        excel_dir = os.path.dirname(excel_file_path)

        # excel file name without extension
        dst_name = os.path.basename(excel_file_path).rsplit('.', 1)[0]

        self.zipped_file_path = os.path.join(excel_dir, dst_name + '.zip')
        self.xlsx_file_path = os.path.join(excel_dir, dst_name + '.xlsx')

        shutil.copy2(excel_file_path, self.zipped_file_path)

        # TODO
        # target_dir += 201402271541 + random string
        self.target_dir = os.path.join(excel_dir, dst_name)
        self.z = ZipFile(self.zipped_file_path)

        try:
            shutil.rmtree(self.target_dir)
        except:
            pass

        os.mkdir(self.target_dir)
        self.z.extractall(self.target_dir)

        self._getSharedString()

    def __del__(self):
        '''remove working directory and file.
        '''
        try:
            # shutil.rmtree(self.target_dir)
            os.remove(self.zipped_file_path)
            os.remove(self.xlsx_file_path)
        except:
            pass

    def _getEtree(self, xml_file_path):
        ''' return etree.XML(xml_file_path)
        '''
        with open(xml_file_path) as f:
            s = f.read()
            return etree.XML(bytes(s, UTF_8))

    def _saveEtree(self, bytes_string, xml_file_path):
        ''' Write bytes string to xml file.
        '''
        with open(xml_file_path, 'w') as f:
            f.write(str(bytes_string, UTF_8))
            return True

    def _makeXMLfilename(self, filename, filepath=None):
        if filepath:
            return os.path.join(filepath, 'sheet') + filename + '.xml'
        return 'sheet' + filename + '.xml'

    def _getSharedString(self):
        ''' get sharedString.xml '''
        self.SharedStringsFile = 'sharedStrings.xml'
        self.SharedStringXML = self._getEtree(os.path.join(self.target_dir, 'xl', self.SharedStringsFile))
        self.sst = self.SharedStringXML
        self.sis = self.SharedStringXML.xpath('//ns:si', namespaces={'ns': NS_SPREADSHEETML})
        '''
        for si in self.sis:
            print(si.find('./ns:t', namespaces={'ns': NS_SPREADSHEETML}).text)
        '''

    @property
    def sheetnames(self):
        '''get sheet names
        '''
        WORKBOOK_NAME = 'workbook.xml'

        # get workbook.xml
        workbook_XML = self._getEtree(os.path.join(self.target_dir, 'xl', WORKBOOK_NAME))

        sheet_elems = workbook_XML.xpath('//ns:sheet', namespaces={'ns': NS_SPREADSHEETML})
        sheetnames = [sheet_elem.get('name') for sheet_elem in sheet_elems]

        return sheetnames

    def _modContentTypes(self, deleteSheetNum=None):
        ''' modify [Content_Types].xml '''
        Content_Types_Name = '[Content_Types].xml'

        # get [Content_Types].xml
        Content_Types_XML = self._getEtree(os.path.join(self.target_dir, Content_Types_Name))
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'

        overrides = Content_Types_XML.xpath('//ns:Override[@ContentType="{}"]'.format(ContentType), namespaces={'ns': NS_CONTENT_TYPES})

        for override in overrides:
            _worksheet = override.get('PartName').rsplit('/', 1)[1]
            _worksheet_num = int(_worksheet[5:-4])  # sheet[NUMBER].xml
            if self.last_worksheet_num < _worksheet_num:
                self.last_worksheet_num = _worksheet_num

        if not deleteSheetNum:
            new_worksheet_num = int(self.last_worksheet_num) + 1
            new_worksheet = self._makeXMLfilename(str(new_worksheet_num))

            new_override = etree.Element('Override')
            new_override.set('PartName', os.path.join('/', 'xl', 'worksheets', new_worksheet))
            new_override.set('ContentType', ContentType)

            # TODO:
            # 기존 worksheet3.xml 노드 뒤에 붙어야 되는데, 더 아래 노드에
            # 들어감.
            # 문제는 없지만 보기 좋게. 이건 중요한 문제는 아님.
            Content_Types_XML.insert(len(overrides), new_override)
            self.last_worksheet_num = new_worksheet_num
        elif deleteSheetNum and not (self.last_worksheet_num == 1):
            deletePartName = os.path.join('/', 'xl', 'worksheets', 'sheet' + str(self.last_worksheet_num) + '.xml')
            deleteElem = Content_Types_XML.find('.//ns:Override[@PartName="{0}"]'.format(deletePartName),\
                                                namespaces={'ns': NS_CONTENT_TYPES})
            deleteElem.getparent().remove(deleteElem)
            self.last_worksheet_num = int(self.last_worksheet_num) - 1
        else:  # deleteSheetNum and self.last_worksheet_num == 1
            print('Excel file must have one sheet at least.')
            # TODO
            # more informative feedback
            raise

        # save [Content_Types].xml
        Content_Types_str = etree.tostring(Content_Types_XML, encoding=UTF_8)
        return self._saveEtree(Content_Types_str, os.path.join(self.target_dir, Content_Types_Name))

    def _modApp(self, sheetname, delete=False):
        ''' modify docProps/app.xml '''
        APP_NAME = 'app.xml'

        # get app.xml
        app_XML = self._getEtree(os.path.join(self.target_dir, 'docProps', APP_NAME))

        # TODO
        # TitleOfParts element can be not exist.
        # In the case, this code will make problem.
        TitleOfParts = app_XML.find('./ns:TitlesOfParts', namespaces={'ns': NS_PROPERTIES})
        TitleOfPartsVector = TitleOfParts.find('./vt:vector', namespaces={'vt': NS_DOC_PROPS_VTYPES})

        i4 = app_XML.find('.//vt:i4', namespaces={'vt': NS_DOC_PROPS_VTYPES})
        lpstrs = TitleOfPartsVector.findall('ns:lpstr', namespaces={'ns': NS_DOC_PROPS_VTYPES})

        # TODO
        # change var name from count to lpstr_count
        count = len(lpstrs)  # current sheets
        if delete:
            count -= 1
        else:
            count += 1
        self.count = str(count)

        # change i4 value
        i4.text = self.count

        for lpstr in lpstrs:
            if sheetname == lpstr.text and delete:
                # change vector size
                new_vector_size = self.count
                TitleOfPartsVector.set('size', new_vector_size)
                # delete the sheet element
                lpstr.getparent().remove(lpstr)
        if not delete:
            # change vector size
            new_vector_size = self.count
            TitleOfPartsVector.set('size', new_vector_size)

            # add new sheet element
            NSMAP = {'vt': NS_DOC_PROPS_VTYPES}
            new_sheet_elem = etree.Element('{%s}lpstr' % NSMAP['vt'], nsmap=NSMAP)
            new_sheet_elem.text = sheetname
            TitleOfPartsVector.insert(len(lpstrs), new_sheet_elem)

        # save app.xml
        app_str = etree.tostring(app_XML, encoding=UTF_8)
        return self._saveEtree(app_str, os.path.join(self.target_dir, 'docProps', APP_NAME))

    def _chkExistWord(self, word):
        ''' check the word is already used. if True, return index.'''
        # TODO
        # 성능 개선 방법
        for index, si in enumerate(self.sis):
            t = si.find('./ns:t', namespaces={'ns': NS_SPREADSHEETML}).text
            # TODO
            # remove this if block
            if t is None:
                print('ERROR: Cannot find <t> element.')
                raise

            if word == t:
                return index
        return False

    def _makeXlsx(self):
        ''' make xlsx
        '''
        shutil.make_archive(self.target_dir, 'zip', self.target_dir)
        os.rename(self.zipped_file_path, self.xlsx_file_path)

        return True

    def _workbook(self, sheetname, new_sheet_name=None, delete=False):
        '''
        modify workbook.xml file
        '''
        workbook_name = 'workbook.xml'

        # get workbook.xml
        workbook_XML = self._getEtree(os.path.join(self.target_dir, 'xl', workbook_name))

        sheets = workbook_XML.find('ns:sheets', namespaces={'ns': NS_SPREADSHEETML})
        if delete:
            deleteId = self._getSheetNum(sheetname)
            deleteElem = workbook_XML.find('.//ns:sheet[@sheetId="{0}"]'.format(deleteId),\
                    namespaces={'ns': NS_SPREADSHEETML})
            deleteElem.getparent().remove(deleteElem)

            sheet_elems = workbook_XML.findall('.//ns:sheet', namespaces={'ns': NS_SPREADSHEETML})
            for el in sheet_elems:
                rId = el.get('{%s}id' % NS_WORKSHEET_R)[3:]
                if int(rId) > int(deleteId):
                    el.set('{%s}id' % NS_WORKSHEET_R, 'rId' + str(int(rId) - 1))
                    el.set('sheetId', str(int(rId) - 1))
        elif new_sheet_name:
            cur_sheetname_node = workbook_XML.find('.//ns:sheet[@name="{}"]'.format(sheetname),\
                    namespaces={'ns': NS_SPREADSHEETML})
            cur_sheetname_node.set('name', new_sheet_name)
        else:
            _sheets = workbook_XML.xpath('//ns:sheet', namespaces={'ns': NS_SPREADSHEETML})
            _Ids = []
            for sheet in _sheets:
                _Ids.append(int(sheet.get('sheetId')))

            new_id = str(max(_Ids) + 1)  # new sheetId
            NSMAP = {'r': NS_WORKSHEET_R}
            new_sheet = etree.Element('sheet', name=sheetname, sheetId=new_id)
            new_sheet.set('{%s}id' % NSMAP['r'], 'rId' + new_id)
            sheets.insert(len(sheets), new_sheet)

        # save workbook.xml
        workbook_str = etree.tostring(workbook_XML, encoding=UTF_8)
        return self._saveEtree(workbook_str, os.path.join(self.target_dir, 'xl', workbook_name))

    def _workbook_refs(self, sheetname=None):
        '''
        modify xl/_refs/workbook.xml.rels file.
        '''
        RELS_NAME = 'workbook.xml.rels'

        # get workbook.xml.rels
        Relationship_XML = self._getEtree(os.path.join(self.target_dir, 'xl', '_rels', RELS_NAME))

        relationships = Relationship_XML.findall('ns:Relationship', namespaces={'ns': NS_RELS})
        rId_str = 'rId'
        rId_len = len(rId_str)

        if sheetname:  # delete sheetname
            dataSheets = Relationship_XML.xpath('//ns:Relationship[@Type="{0}"]'.format(NS_REL_WORKSHEET), namespaces={'ns': NS_RELS})

            target = self._makeXMLfilename(str(len(dataSheets)), 'worksheets')

            deleteElem = Relationship_XML.find('.//ns:Relationship[@Target="{0}"]'.format(target), namespaces={'ns': NS_RELS})
            # deleteElem.decompose()
            deleteElem.getparent().remove(deleteElem)

            relationships = Relationship_XML.findall('ns:Relationship', namespaces={'ns': NS_RELS})
            for el in relationships:
                _id = int(el.get('Id')[len(rId_str):])
                if _id > len(dataSheets):
                    el.set('Id', 'rId' + str(_id - 1))
        else:
            rIds = [int(relationship.get('Id')[rId_len:]) for relationship in relationships]
            '''
            뭐하는 코드인지 모르겠음.
            if relationship['Id'][rID_len:] == self.count:
                print(relationship['Id'][rID_len:])
                break
            '''
            rIds = sorted(rIds, reverse=True)
            changeIds = list(filter(lambda rId: rId >= int(self.count), rIds))

            # change ID in <Relationship@Id> >= number of sheets
            for rId in changeIds:
                Id = rId_str + str(rId + 1)
                _elem = Relationship_XML.find('ns:Relationship[@Id="rId{0}"]'.format(rId),
                                              namespaces={'ns': NS_RELS})
                _elem.set('Id', Id)

            new_rId = 'rId' + self.count
            Target = "worksheets/sheet{0}.xml".format(self.count)
            Target = self._makeXMLfilename(self.count, 'worksheets')

            new_relationship = etree.Element('Relationship', Id=new_rId, Type=NS_REL_WORKSHEET, Target=Target)

            _relationship = Relationship_XML.xpath('//ns:Relationships', namespaces={'ns': NS_RELS})[0]
            _relationship.insert(len(relationships), new_relationship)

        # save workbook.xml.rels
        Relationship_str = etree.tostring(Relationship_XML, encoding=UTF_8)
        return self._saveEtree(Relationship_str, os.path.join(self.target_dir, 'xl', '_rels', RELS_NAME))

    def _makeSheet(self):
        '''
        copy template file to xl/worksheets/sheet[n].xml
        '''
        template = os.path.join(os.path.dirname(__file__), 'templates', 'xl', 'worksheets', 'sheet.xml')
        new_sheet_file_name = self._makeXMLfilename(self.count)
        target = os.path.join(self.target_dir, 'xl', 'worksheets', new_sheet_file_name)
        shutil.copy2(template, target)

    def _getSheetNum(self, sheetname):
        '''
        Get sheet id number.
        '''
        if sheetname not in self.sheetnames:
            print('Sheet name [{0}] does not exist'.format(sheetname))
            # TODO
            # more informative feedback.
            raise
        try:
            WORKBOOK_NAME = 'workbook.xml'

            workbook_XML = self._getEtree(os.path.join(self.target_dir, 'xl', WORKBOOK_NAME))

            sheet = workbook_XML.find('.//ns:sheet[@name="{}"]'.format(sheetname), namespaces={'ns': NS_SPREADSHEETML})
            return sheet.get('sheetId')
        except:
            print('ERROR: Cannot get sheet number from workbook.xml')
            raise

    def _copySheet(self, orig_sheetname, copy_name):
        # TODO
        # 시트 복사 할 때 기존 시트에 있는 문자열 개수를 가져와서
        # 전체 count 값에 더해 주어야 함.
        # print(self.sst.get('count'))
        # count = str(int(self.sst.get('count')) * 2)
        # self.sst.set('count', count)
        # print(self.sst.get('count'))

        sheetnum = self._getSheetNum(orig_sheetname)
        file_name = self._makeXMLfilename(str(sheetnum))
        new_file_name = self._makeXMLfilename(self.count)
        orig_sheet_file = os.path.join(self.target_dir, 'xl', 'worksheets', file_name)
        new_sheet_file = os.path.join(self.target_dir, 'xl', 'worksheets', new_file_name)
        shutil.copy2(orig_sheet_file, new_sheet_file)

    def _step_uniqueCount(self, step=1):
        new_uniqueCount = int(self.sst.get('uniqueCount')) + step
        self.sst.set('uniqueCount', str(new_uniqueCount))

    @property
    def _get_count_sharedStrings(self):
        aaa = int(self.sst.get('count'))
        return aaa

    def _set_count_sharedStrings(self, add_count_value):
        new_count_value = str(self._get_count_sharedStrings + add_count_value)
        self.sst.set('count', new_count_value)
        self._saveEtree(etree.tostring(self.SharedStringXML),\
            os.path.join(self.target_dir, 'xl', 'sharedStrings.xml'))

    def _get_length_type_is_s_sheetfile(self, sheetname):
        '''
        get the length of <c> element that t property is 's'\
        in xl/worksheets/sheet[#].xml file.
        '''
        sheetnum = self._getSheetNum(sheetname)
        file_name = self._makeXMLfilename(str(sheetnum))
        sheet_file = os.path.join(self.target_dir, 'xl', 'worksheets', file_name)
        sheetXML = self._getEtree(sheet_file)
        return len(sheetXML.xpath('//ns:c[@t="s"]', namespaces={'ns': NS_SPREADSHEETML}))

    def addSheet(self, new_sheet_name):
        # TODO:
        # select sheetId
        # default: last
        # select: index number
        self._modContentTypes()
        self._modApp(new_sheet_name)

        self._workbook(new_sheet_name)
        self._workbook_refs()
        self._makeSheet()

    def copySheet(self, orig_sheetname, copy_name=None):
        if copy_name is None:
            copy_name = orig_sheetname + ' copy'

        curr_s_type_count = self._get_length_type_is_s_sheetfile(orig_sheetname)
        self._set_count_sharedStrings(curr_s_type_count)

        self._modContentTypes()
        if self._modApp(copy_name):
            # TODO
            # save()메소드 호출하지 않는 방법?
            self.save()
        else:
            # _modApp()이 실패하면 _modApp()에서 오류 발생하도록.
            print('Failed to copy. The sheet name is already used. Use another name.')
            return False

        self._workbook(copy_name)
        self._workbook_refs()
        self._copySheet(orig_sheetname, copy_name)

    def deleteSheet(self, sheetname):
        # TODO
        # decrease count property of <sst> in 'sharedString.xml' file.
        curr_s_type_count = self._get_length_type_is_s_sheetfile(sheetname)
        self._set_count_sharedStrings(curr_s_type_count * -1)

        self._del_sheet_id = self._getSheetNum(sheetname)
        sheetidx = self._makeXMLfilename(self._del_sheet_id)
        self._modContentTypes(sheetidx)
        self._modApp(sheetname, delete=True)
        self._workbook_refs(sheetname)  # This func must be called before _workbook()
        self._workbook(sheetname, delete=True)

        # TODO:
        # xl/worksheets/sheet[#].xml file numbering.
        worksheets_path = os.path.join(self.target_dir, 'xl', 'worksheets')
        target = os.path.join(worksheets_path, sheetidx)
        os.remove(target)
        pattern_sheets = 'sheet*.xml'

        xmls = fnmatch.filter(os.listdir(worksheets_path), pattern_sheets)
        for xml in xmls:
            num = int(xml[5:-4])
            if num > int(self._del_sheet_id):
                new_name = self._makeXMLfilename(str(num - 1))
                os.rename(os.path.join(worksheets_path, xml), os.path.join(worksheets_path, new_name))

    def renameSheet(self, orig_sheetname, new_sheet_name):
        self._workbook(orig_sheetname, new_sheet_name)

    def _saveSharedString(self):
        ''' save sharedString.xml '''
        self._saveEtree(etree.tostring(self.SharedStringXML, encoding='utf-8'),\
                        os.path.join(self.target_dir, 'xl', self.SharedStringsFile))

    def _set_v_text(self, cell, index):
        v = cell.find('./ns:v', namespaces={'ns': NS_SPREADSHEETML})
        v.text = str(index)

    def edit(self, sheetname, changed_data_json):
        # TODO
        # 단순히 raise 말고 적당한 피드백을 주도록.
        # use decorator?
        if sheetname not in self.sheetnames:
            print('Sheet name [{0}] does not exist'.format(sheetname))
            raise

        sheet_number = self._getSheetNum(sheetname)

        self.changed = json.loads(changed_data_json)

        ''' get sheet.xml '''
        worksheet_dir = os.path.join(self.target_dir, 'xl', 'worksheets')

        # xml_file = os.path.join(self.worksheet_dir, self.sheet_number)
        xml_file = self._makeXMLfilename(sheet_number, worksheet_dir)

        # TODO
        # To use self._getEtree() method
        # self.sheetXML = self._getEtree(xml_file)
        self.sheetXML = etree.parse(xml_file).getroot()

        ''' change value '''
        cells = self.changed.keys()
        for _cell in cells:
            cell = self.sheetXML.find('.//ns:c[@r="{0}"]'.format(_cell), namespaces={'ns': NS_SPREADSHEETML})
            if cell is None:  # nothing in cell before.
                r = re.findall(r'\d+', _cell)[0]
                row = self.sheetXML.xpath('//ns:row[@r="{}"]'.format(r), namespaces={'ns': NS_SPREADSHEETML})
                if len(row) is 0:
                    NSMAP = {'ns': NS_SPREADSHEETML}
                    new_row = etree.Element('{%s}row' % NSMAP['ns'], nsmap=NSMAP, r=r)
                    sheetData = self.sheetXML.find('ns:sheetData', namespaces={'ns': NS_SPREADSHEETML})
                    sheetData.append(new_row)
                    row = sheetData.find('.//ns:row[@r="{}"]'.format(r), namespaces={'ns': NS_SPREADSHEETML})
                new_v = etree.Element('v')
                # count = str(int(self.sst.get('count')) + 1)
                # self.sst.set('count', count)

                if isinstance(self.changed[_cell], int):
                    # new_c = self.sheetXML.new_tag('c', r=_cell, s='1')
                    # new_c maybe require namespaces.
                    new_c = etree.Element('c', r=_cell, s='1')
                    new_v.text = str(self.changed[_cell])
                    new_c.append(new_v)
                    row.append(new_c)
                else:  #Unicode String
                    # TODO
                    # Add s attribute if necessary.
                    new_c = etree.Element('{%s}c' % NSMAP['ns'], r=_cell, t='s')
                    new_v = etree.Element('{%s}v' % NSMAP['ns'])
                    new_c.insert(0, new_v)
                    index = self._chkExistWord(self.changed[_cell])
                    if index:  # already used word
                        self._set_v_text(new_c, index)
                        # TODO
                        # count가 너무 많이 더해짐.
                    else:  # new word
                        self._step_uniqueCount(step=1)
                        # create a <si><t>text</t></si> element
                        new_si = etree.Element('{%s}si' % NSMAP['ns'])
                        new_t = etree.Element('{%s}t' % NSMAP['ns'])
                        new_t.text = self.changed[_cell]
                        new_si.insert(0, new_t)
                        sis = self.SharedStringXML.findall('ns:si', namespaces={'ns': NS_SPREADSHEETML})
                        self.sst.insert(len(sis), new_si)
                        sis = self.SharedStringXML.findall('ns:si', namespaces={'ns': NS_SPREADSHEETML})
                        new_v_num = len(sis)
                        new_v.text = str(new_v_num - 1)  # last si index
                        # TODO
                        # xml 저장할 때, 셀 값을 utf-8로 인코딩해서 저장하도록?
                        self._saveEtree(etree.tostring(self.SharedStringXML),\
                            os.path.join(self.target_dir, 'xl', 'sharedStrings.xml'))

                    count = str(int(self.sst.get('count')) + 1)
                    self.sst.set('count', count)

                    ### add new element.
                    new_c.append(new_v)
                    row.append(new_c)
                    ###
            else:  # something in cell before.
                v = cell.find('ns:v', namespaces={'ns': NS_SPREADSHEETML})
                if isinstance(self.changed[_cell], int):  # new value is integer.
                    try:
                        cell.attrib.pop('t')  # remove 't' attribute if already exist
                    except:
                        pass
                    v.text = str(self.changed[_cell])
                else:  # new value is not integer. it must be unicode.
                    # TODO
                    # 기존 문자열 값이 새 문자열로 바뀔 때
                    # count -= 1을 해야하나?
                    # 기존 문자열이 더 이상 쓰이지 않을 경우
                    # uniqueCount -=1을 해야하나?
                    try:
                        cell.set('t', 's')
                    except:
                        pass
                    if v is None:
                        new_v = etree.Element('v')
                        cell.insert(0, new_v)
                    else:
                        pass  # TODO: remove this else block.

                    child = cell.find('ns:v', namespaces={'ns': NS_SPREADSHEETML})
                    if child is not None:
                        index = self._chkExistWord(self.changed[_cell])
                        if index:  # already used word
                            self._set_v_text(cell, index)
                        else:  # new word
                            self._step_uniqueCount(step=1)
                            # create a <si><t>text</t></si> element
                            NSMAP = {'ns': NS_SPREADSHEETML}
                            new_si = etree.Element('{%s}si' % NSMAP['ns'], nsmap=NSMAP)
                            new_t = etree.Element('t')
                            new_t.text = self.changed[_cell]
                            len_sis = len(self.SharedStringXML.findall('ns:si', namespaces={'ns': NS_SPREADSHEETML}))
                            new_si.insert(0, new_t)
                            self.SharedStringXML.insert(len_sis, new_si)
                            v = cell.find('ns:v', namespaces={'ns': NS_SPREADSHEETML})
                            if v is not None:
                                v.text = str(len_sis)  # last si index
                            self._saveEtree(etree.tostring(self.SharedStringXML),\
                                os.path.join(self.target_dir, 'xl', 'sharedStrings.xml'))
                    count = str(int(self.sst.get('count')) + 1)
                    self.sst.set('count', count)
        else:
            ''' save modified sheet file(xml) '''
        # TODO
        # _saveSharedString()을 edit()할 때가 아니라
        # save()할 때 호출하도록.
        # edit()할 때마다 sharedString.xml 파일을 수정하는 것은
        # 너무 비효율적임.
        self._saveSharedString()
        return self._saveEtree(etree.tostring(self.sheetXML),\
                self._makeXMLfilename(sheet_number, worksheet_dir))

    def save(self):
        self._makeXlsx()
