import copy
import json
import statistics
import subprocess
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
RAW=['测试ID','题目来源','题号','耗时 时分秒','LLM完成调用数 次','输入Token 万','输出Token 万','缓存命中率 %','工具调用 次','工具错误 次','会话ID','用户消息ID','开始时间','结束时间']
def letters(n):
    text=''
    while n:n,v=divmod(n-1,26);text=chr(65+v)+text
    return text


def write_workbook(path,rows,raw=None):
    def sheet(rows):
        root=ET.Element('worksheet',xmlns=a.NS['s']);data=ET.SubElement(root,'sheetData')
        for i,row in enumerate(rows,1):
            r=ET.SubElement(data,'row',r=str(i))
            for j,value in enumerate(row,1):
                cell=ET.SubElement(r,'c',r=letters(j)+str(i))
                if isinstance(value,str):cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,'is'),'t').text=value
                elif value is not None:ET.SubElement(cell,'v').text=str(value)
        return ET.tostring(root,encoding='utf-8')
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        raw_sheet='<sheet name="逐次记录" sheetId="2" r:id="r2"/>' if raw is not None else ''
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{a.NS["s"]}" xmlns:r="{a.R}"><sheets><sheet name="任务均值" sheetId="1" r:id="r1"/>{raw_sheet}</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="worksheets/sheet2.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml',sheet(rows))
        if raw is not None:z.writestr('xl/worksheets/sheet2.xml',sheet(raw))


def workbook(path,name,seconds=(1,3,10),drop_cache=False,extra=False):
    rows=[[name+' 端侧日志统计'],HEADERS,
        ['综合能力','1-1','Question A',2,(seconds[0]+seconds[1])/2/86400,2,100/10000,10/10000,.5,4,0],
        ['AlphaData','1-1','Question B',1,seconds[2]/86400,4,200/10000,20/10000,None if drop_cache else .9,8,2]]
    if extra:rows.append(['DeepInsight','2-1','Extra',1,1/86400,1,1,1,.2,2,1])
    raw=[RAW]
    for i,s in enumerate(seconds):
        raw.append([f'run-{i}','综合能力' if i<2 else 'AlphaData','1-1',s/86400,2 if i<2 else 4,
            (100 if i<2 else 200)/10000,(10 if i<2 else 20)/10000,.5 if i<2 else None if drop_cache else .9,4 if i<2 else 8,0 if i<2 else 2,
            name+f'-session-{i}',name+f'-message-{i}',f'2031-01-01T00:00:{i:02d}',None])
    if extra:raw.append(['extra','DeepInsight','2-1',1/86400,1,1,1,.2,2,1,name+'-extra-session',name+'-extra-message',None,None])
    write_workbook(path,rows,raw)


def fragment(path,original,indices,change=None,omit_raw=False):
    """Build independently recomputed task summaries for selected fixture runs."""
    sheets,_=a.read_xlsx(original)
    cells=lambda row:[row.get(letters(i),{}).get('value') for i in range(1,len(RAW)+1)]
    raw=[cells(sheets['逐次记录'][i+1]) for i in indices]
    if change:
        for i,row in enumerate(raw):change(i,row)
    titles={(r['A']['value'],r['B']['value']):r['C']['value'] for r in sheets['任务均值'][2:]}
    rows=[[sheets['任务均值'][0]['A']['value']],HEADERS]
    for key in sorted({(r[1],r[2]) for r in raw}):
        values=[r for r in raw if (r[1],r[2])==key]
        means=[]
        for column in range(3,10):
            known=[r[column] for r in values if r[column] is not None]
            means.append(statistics.mean(known) if known else None)
        rows.append([*key,titles[key],len(values),*means])
    write_workbook(path,rows,None if omit_raw else [RAW,*raw])


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

    def test_distributed_datasets_match_whole_workbooks_after_union(self):
        first,second=self.inputs(extra=True);expected=a.compare([first,second]);files=[]
        for original in (first,second):
            for batch,indices in enumerate(([0,1,2],[3]) if original is first else ([0,1],[2])):
                path=self.root/f'{original["name"]}-{batch}.xlsx';fragment(path,original['path'],indices)
                files.append(a.load_workbook(path))
        actual=a.compare(files)
        self.assertEqual(actual['common_tasks'],expected['common_tasks'])
        for x,y in zip(actual['objects'],expected['objects']):
            self.assertEqual(x['means'],y['means']);self.assertEqual(x['run_count'],y['run_count'])
            self.assertEqual(len(x['sources']),2)

    def test_same_local_test_id_is_distinct_execution_and_means_are_run_weighted_per_task(self):
        first,second=self.inputs();p=self.root/'other-worker.xlsx'
        def change(i,row):
            row[0]='run-0';row[3]=9/86400;row[5]=400/10000
            row[10]='new-session';row[11]='new-message'
        fragment(p,first['path'],[0],change)
        data=a.compare([first,a.load_workbook(p),second]);o=data['objects'][0]
        self.assertEqual(o['run_count'],4);self.assertEqual(o['time_n'],4)
        self.assertEqual(o['means']['input_tokens'],200)
        self.assertAlmostEqual(o['means']['seconds'],(13/3+10)/2)
        self.assertAlmostEqual(o['means']['p95'],9.85)
        self.assertEqual(o['merge']['deduplicated_run_rows'],0)

    def test_copied_file_and_overlapping_exports_not_double_counted(self):
        first,second=self.inputs();copy_path=self.root/'copy.xlsx';copy_path.write_bytes(Path(first['path']).read_bytes())
        p=self.root/'subset.xlsx';fragment(p,first['path'],[0,2],lambda i,row:row.__setitem__(0,f'renumbered-{i}'))
        expected=a.compare([first,second])['objects'][0]
        result=a.compare([first,a.load_workbook(copy_path),a.load_workbook(p),second])['objects'][0]
        self.assertEqual(result['means'],expected['means']);self.assertEqual(result['run_count'],3)
        self.assertEqual(result['merge']['deduplicated_run_rows'],2)
        self.assertEqual(len(result['merge']['duplicate_files']),1)
        self.assertEqual(sum(len(r['evidence']) for r in result['runs']),5)

    def test_overlap_missing_value_filled_without_zero_or_losing_real_zero(self):
        first,second=self.inputs(drop_cache=True);p=self.root/'more-complete.xlsx'
        fragment(p,first['path'],[2],lambda i,row:row.__setitem__(7,0))
        obj=a.compare([first,a.load_workbook(p),second])['objects'][0]
        self.assertEqual(obj['means']['cache_rate'],.25);self.assertEqual(obj['run_count'],3)

    def test_overlap_conflicts_stop_instead_of_picking_one(self):
        first,second=self.inputs();p=self.root/'conflict.xlsx'
        fragment(p,first['path'],[0],lambda i,row:row.__setitem__(5,.04))
        with self.assertRaisesRegex(a.InputError,'指标冲突'):a.compare([first,a.load_workbook(p),second])
        fragment(p,first['path'],[0],lambda i,row:row.__setitem__(12,'2031-02-01T00:00:00'))
        with self.assertRaisesRegex(a.InputError,'时间冲突'):a.compare([first,a.load_workbook(p),second])

    def test_overlap_without_native_identity_or_raw_stops_but_disjoint_summary_allowed(self):
        first,second=self.inputs();p=self.root/'no-id.xlsx'
        fragment(p,first['path'],[0],lambda i,row:row.__setitem__(10,None))
        with self.assertRaisesRegex(a.InputError,'会话ID和用户消息ID'):a.compare([first,a.load_workbook(p),second])
        fragment(p,first['path'],[0],omit_raw=True)
        with self.assertRaisesRegex(a.InputError,'完整逐次记录'):a.compare([first,a.load_workbook(p),second])
        fragment(p,first['path'],[2],omit_raw=True);q=self.root/'other-task.xlsx';fragment(q,first['path'],[0,1])
        obj=a.compare([a.load_workbook(p),a.load_workbook(q),second])['objects'][0]
        self.assertEqual(obj['run_count'],3);self.assertEqual(obj['time_n'],2)
        self.assertEqual(obj['means']['input_tokens'],150)

    def test_recursive_cli_discovers_same_names_in_person_subdirectories(self):
        expected=list(self.inputs());inputs=self.root/'inputs'
        for obj in expected:
            for person,indices in [('person1',[0,1]),('person2',[2])]:
                folder=inputs/person;folder.mkdir(parents=True,exist_ok=True)
                fragment(folder/(obj['name']+'_过程指标统计.xlsx'),obj['path'],indices)
        (inputs/'~$A_过程指标统计.xlsx').write_bytes(b'not a workbook')
        (inputs/'unrelated.xlsx').write_bytes(b'not a workbook')
        target=self.root/'result.json'
        subprocess.run([sys.executable,str(ROOT/'skills/cost-statistics/scripts/aggregate.py'),
            '--input-dir',str(inputs),'--output',str(target)],capture_output=True,check=True)
        data=json.loads(target.read_text(encoding='utf8'))
        self.assertEqual([o['means'] for o in data['objects']],[o['means'] for o in a.compare(expected)['objects']])
        self.assertTrue(all(len(o['sources'])==2 for o in data['objects']))

    def test_repeated_object_name_aliases_and_file_title_precede_filename(self):
        first,second=self.inputs();p=self.root/'renamed.xlsx';fragment(p,first['path'],[2])
        self.assertEqual(a.load_workbook(p)['name'],'A')
        first['name']='mobilework';other=a.load_workbook(p,'MobileWork')
        result=a.compare([first,other,second]);self.assertEqual(result['objects'][0]['name'],'MobileWork')
        self.assertEqual(result['objects'][0]['run_count'],3)

    def test_identity_cannot_reference_two_questions_or_repeat_within_file(self):
        first,second=self.inputs();third=copy.deepcopy(first);third['sha256']='new';third['path']='another.xlsx'
        third['raw'][2]['session_id']=third['raw'][0]['session_id'];third['raw'][2]['user_message_id']=third['raw'][0]['user_message_id']
        with self.assertRaisesRegex(a.InputError,'不同题目|同一工作簿'):a.compare([first,third,second])
        first['raw'][1]['session_id']=first['raw'][0]['session_id'];first['raw'][1]['user_message_id']=first['raw'][0]['user_message_id']
        with self.assertRaisesRegex(a.InputError,'同一工作簿'):a.compare([first,second])

    def test_same_object_conflicting_task_titles_rejected(self):
        first,second=self.inputs();p=self.root/'overlap.xlsx';fragment(p,first['path'],[0])
        other=a.load_workbook(p);other['tasks'][('综合能力','1-1')]['title']='Different'
        with self.assertRaisesRegex(a.InputError,'题目名称冲突'):a.compare([first,other,second])


if __name__=='__main__':unittest.main()
