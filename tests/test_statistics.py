import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile,ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'skills/cost-statistics/scripts'))
import aggregate as a


HEADERS=['题目来源','题号','题目','独立次数','耗时 时分秒','LLM完成调用数 次','输入Token 万','输出Token 万','缓存命中率 %','工具调用 次','工具错误 次']
RAW=['测试ID','题目来源','题号','耗时 时分秒','LLM完成调用数 次','输入Token 万','输出Token 万','缓存命中率 %','工具调用 次','工具错误 次']
def letters(n):
    text=''
    while n:n,v=divmod(n-1,26);text=chr(65+v)+text
    return text


def workbook(path,name,seconds=(1,3,10),drop_cache=False,extra=False):
    def sheet(rows):
        root=ET.Element('worksheet',xmlns=a.NS['s']);data=ET.SubElement(root,'sheetData')
        for i,row in enumerate(rows,1):
            r=ET.SubElement(data,'row',r=str(i))
            for j,value in enumerate(row,1):
                cell=ET.SubElement(r,'c',r=letters(j)+str(i))
                if isinstance(value,str):cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,'is'),'t').text=value
                elif value is not None:ET.SubElement(cell,'v').text=str(value)
        return ET.tostring(root,encoding='utf-8')
    rows=[[name+' 端侧日志统计'],HEADERS,
        ['综合能力','1-1','Question A',2,(seconds[0]+seconds[1])/2/86400,2,100/10000,10/10000,.5,4,0],
        ['AlphaData','1-1','Question B',1,seconds[2]/86400,4,200/10000,20/10000,None if drop_cache else .9,8,2]]
    if extra:rows.append(['DeepInsight','2-1','Extra',1,1/86400,1,1,1,.2,2,1])
    raw=[RAW]
    for i,s in enumerate(seconds):
        raw.append([f'run-{i}','综合能力' if i<2 else 'AlphaData','1-1',s/86400,2 if i<2 else 4,
            (100 if i<2 else 200)/10000,(10 if i<2 else 20)/10000,.5 if i<2 else None if drop_cache else .9,4 if i<2 else 8,0 if i<2 else 2])
    if extra:raw.append(['extra','DeepInsight','2-1',1/86400,1,1,1,.2,2,1])
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{a.NS["s"]}" xmlns:r="{a.R}"><sheets><sheet name="任务均值" sheetId="1" r:id="r1"/><sheet name="逐次记录" sheetId="2" r:id="r2"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="worksheets/sheet2.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml',sheet(rows));z.writestr('xl/worksheets/sheet2.xml',sheet(raw))


class Tests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def inputs(self,**kwargs):
        p=self.root/'a.xlsx';q=self.root/'b.xlsx';workbook(p,'A',**kwargs);workbook(q,'B',(2,4,8))
        return a.load_workbook(p),a.load_workbook(q)

    def test_task_equal_means_and_raw_percentiles_with_repeats(self):
        first,second=self.inputs();data=a.compare([first,second]);o=data['objects'][0]
        self.assertEqual(o['means']['input_tokens'],150)
        self.assertAlmostEqual(o['means']['seconds'],6)
        self.assertAlmostEqual(o['means']['p50'],3)
        self.assertAlmostEqual(o['means']['p95'],9.3)
        self.assertEqual(o['run_count'],3);self.assertEqual(o['time_n'],3)
        self.assertEqual(len(data['common_tasks']),2)

    def test_same_number_in_different_sources_not_merged(self):
        data=a.compare(list(self.inputs()))
        self.assertEqual({t['source'] for t in data['common_tasks']},{'综合能力','AlphaData'})

    def test_different_coverage_uses_intersection(self):
        data=a.compare(list(self.inputs(extra=True)))
        self.assertEqual(len(data['common_tasks']),2);self.assertEqual(data['objects'][0]['run_count'],3)
        self.assertEqual(data['objects'][0]['excluded_tasks'],[['DeepInsight','2-1']])

    def test_missing_metric_uses_shared_valid_set_and_keeps_zero(self):
        data=a.compare(list(self.inputs(drop_cache=True)))
        self.assertEqual(data['metric_task_counts']['cache_rate'],1)
        self.assertEqual(data['objects'][0]['means']['cache_rate'],.5)
        self.assertEqual(data['objects'][1]['means']['cache_rate'],.5)
        self.assertEqual(data['objects'][0]['means']['tool_errors'],1)

    def test_zero_duration_and_percentile_boundary(self):
        self.assertEqual(a.percentile([0,0,0],.95),0)
        self.assertEqual(a.percentile([5],.95),5)
        self.assertIsNone(a.percentile([],.5))

    def test_conflicting_names_and_duplicate_inputs_rejected(self):
        first,second=self.inputs();second['tasks'][('综合能力','1-1')]['title']='Different'
        with self.assertRaises(a.InputError):a.compare([first,second])
        with self.assertRaises(a.InputError):a.compare([first,copy.deepcopy(first)])

    def test_missing_raw_does_not_estimate_percentiles(self):
        first,second=self.inputs();first['raw']=[]
        data=a.compare([first,second]);self.assertIsNone(data['objects'][0]['means']['p95'])

    def test_many_tasks_have_no_fixed_sample_limit(self):
        first,second=self.inputs()
        for obj in (first,second):
            template=obj['tasks'][('综合能力','1-1')]
            obj['tasks']={(f'Dataset-{i%3}',f'1-{i+1}'):copy.deepcopy(template) for i in range(150)}
            obj['raw']=[dict(key=key,run_id=f'{i}-{j}',date=None,metrics={**t['metrics'],'seconds':float(j+1)})
                for i,(key,t) in enumerate(obj['tasks'].items()) for j in range(2)]
        data=a.compare([first,second]);self.assertEqual(len(data['common_tasks']),150)
        self.assertEqual(data['objects'][0]['run_count'],300)
        self.assertEqual(data['objects'][0]['time_n'],300)
        self.assertEqual(data['objects'][0]['means']['p50'],1.5)

    def test_source_workbooks_unchanged(self):
        first,second=self.inputs();before={p:p.read_bytes() for p in self.root.glob('*.xlsx')}
        a.compare([first,second])
        self.assertTrue(all(p.read_bytes()==raw for p,raw in before.items()))

    def test_formula_without_cached_value_stops(self):
        cell=dict(value=None,formula=True,evidence='book#任务均值!A7')
        with self.assertRaises(a.InputError):a.get_metric({'A':cell},{'输入Token万':'A'},'input_tokens',[])


if __name__=='__main__':unittest.main()
