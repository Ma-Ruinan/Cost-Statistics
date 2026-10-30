import copy
from pathlib import Path
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile
import test_statistics as fixtures
import aggregate
import report


class ReportTests(unittest.TestCase):
    def test_report_has_only_requested_metrics_and_no_fixed_batch_content(self):
        try:font=report.font_path()
        except aggregate.InputError:self.skipTest('No local CJK font for rendering fixture')
        with tempfile.TemporaryDirectory() as root:
            root=Path(root)
            for name in ('Object A','Object B'):fixtures.workbook(root/(name+'.xlsx'),name)
            data=aggregate.compare([aggregate.load_workbook(p) for p in sorted(root.glob('*.xlsx'))])
            path=report.build(data,root/'report.docx',date='2031-05-04',font=font)
            with ZipFile(path) as z:
                xml=ET.fromstring(z.read('word/document.xml'));ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                text=''.join(t.text or '' for t in xml.findall('.//w:t',ns))
                self.assertEqual(len(xml.findall('.//w:tbl',ns)),2)
                self.assertIn('平均工具错误次数',text);self.assertIn('端到端耗时P95',text)
                self.assertIn('Object A',text);self.assertIn('2031-05-04',text)
                self.assertNotIn('搜索',text);self.assertNotIn('数据来源',text);self.assertNotIn('AI Agent',text)
                self.assertNotIn('MobileWork',text)

    def test_four_objects_and_large_title_numbers_do_not_change_analysis_arithmetic(self):
        with tempfile.TemporaryDirectory() as root:
            root=Path(root);items=[]
            for i in range(4):
                path=root/f'{i}.xlsx';fixtures.workbook(path,f'Object {i}',seconds=(1+i,3+i,10+i));items.append(aggregate.load_workbook(path))
            data=aggregate.compare(items);notes=report.suggested_analysis(data)
            self.assertEqual(len(data['objects']),4)
            self.assertTrue(notes);self.assertTrue(all('heading' in n and 'text' in n for n in notes))
            self.assertNotIn('MobileWork',str(notes))


if __name__=='__main__':unittest.main()
