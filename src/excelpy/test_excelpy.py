import os
import shutil
import lxml
from lxml import etree

from excelpy import ExcelPy
import unittest

UTF_8 = 'utf-8'

from xlsx_ns import NS_CONTENT_TYPES


class ExcelPyTest(unittest.TestCase):
    def setUp(self):
        try:
            shutil.rmtree('test_slayers')
            os.remove('test_slayers.zip')
            os.remove('test_slayers.xlsx')
        except:
            pass

    def tearDown(self):
        self.assertFalse(os.path.exists('test_slayers'))
        self.assertFalse(os.path.exists('test_slayers.zip'))

    def test_getETree(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        xml_file = os.path.join('test_slayers', '[Content_Types].xml')
        self.assertEqual(lxml.etree._Element, type(_excel._getEtree(xml_file)))

    def test_saveETree(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        xml_file = os.path.join('test_slayers', '[Content_Types].xml')
        Content_Types_XML = _excel._getEtree(xml_file)
        Content_Types_str = etree.tostring(Content_Types_XML, encoding=UTF_8)
        self.assertTrue(_excel._saveEtree(Content_Types_str, xml_file))

    def test_get_sheet_names(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        self.assertEqual(['Slayers', 'Sheet2', 'Sheet3'], _excel.sheetnames)

    def test_modApp(self):
        pass

    def test_modContentType(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        Content_Types_XML = _excel._getEtree(os.path.join(_excel.target_dir, '[Content_Types].xml'))
        overrides = Content_Types_XML.xpath('//ns:Override', namespaces={'ns': NS_CONTENT_TYPES})
        for override in overrides:
            if 'worksheets' in override.get('PartName'):
                last_worksheet = override.get('PartName').rsplit('/', 1)[1]
        last_worksheet_num = last_worksheet[5:-4]  # sheet[NUMBER].xml

        self.assertIsInstance(Content_Types_XML, lxml.etree._Element)
        self.assertNotEqual(Content_Types_XML, None)
        self.assertIsInstance(overrides, list)
        self.assertIsNot(len(overrides), 0)
        self.assertEqual(last_worksheet_num, '3')

    def test_add_sheet(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.addSheet('슬레이어즈')
        self.assertEqual(['Slayers', 'Sheet2', 'Sheet3', '슬레이어즈'], _excel.sheetnames)

    def test_delete_sheet(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.deleteSheet('Sheet2')
        self.assertEqual(['Slayers', 'Sheet3'], _excel.sheetnames)
        self.assertNotEqual(['Slayers', 'Sheet2', 'Sheet3'], _excel.sheetnames)

    def test_copy_sheet_without_new_sheetname(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.copySheet('Sheet2')
        self.assertEqual(['Slayers', 'Sheet2', 'Sheet3', 'Sheet2 copy'], _excel.sheetnames)
        self.assertNotEqual(['Slayers', 'Sheet2', 'Sheet3', 'Sheet2 Copy'], _excel.sheetnames)

    def test_copy_sheet_with_new_sheetname(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.copySheet('Sheet2', 'Fantasy')
        self.assertEqual(['Slayers', 'Sheet2', 'Sheet3', 'Fantasy'], _excel.sheetnames)
        self.assertNotEqual(['Slayers', 'Sheet2', 'Sheet3'], _excel.sheetnames)

    def test_rename_sheetname(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.renameSheet('Sheet2', 'Fantasy')
        self.assertEqual(['Slayers', 'Fantasy', 'Sheet3'], _excel.sheetnames)
        self.assertNotEqual(['Slayers', 'Sheet2', 'Sheet3'], _excel.sheetnames)

    def test_complex_test(self):
        test_excel = shutil.copyfile('slayers.xlsx', 'test_slayers.xlsx')
        _excel = ExcelPy(test_excel)
        _excel.copySheet('Slayers', '슬레이어즈')
        _excel.copySheet('Sheet2')
        _excel.renameSheet('Sheet2', 'Fantasy')
        _excel.addSheet('BrownEyed')
        _excel.deleteSheet('Sheet3')
        _excel.addSheet('Soul')
        self.assertEqual(['Slayers', 'Fantasy', '슬레이어즈', 'Sheet2 copy', 'BrownEyed', 'Soul'], _excel.sheetnames)
        self.assertNotEqual(['Slayers', 'Sheet2', '슬레이어즈', 'Sheet2 copy', 'BrownEyed'], _excel.sheetnames)

    def test_si_index(self):
        '''
        used word's index of <sis> in 'sharedStrings.xml'
        '''
        pass

    def test_get_shared_string_count(self):
        '''
        get the value of count property of <sst> in xl/sharedString.xml
        '''
        pass

    def test_get_current_sheet_words_count(self):
        '''
        get the value of count of current sheet.
        '''
        pass


if __name__ == '__main__':
    unittest.main()
