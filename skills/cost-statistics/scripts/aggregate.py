"""Read accounting workbooks without Office, modifying neither files nor formulas."""
import argparse
import collections
import datetime as dt
import hashlib
import json
import math
from pathlib import Path
import posixpath
import re
import statistics
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
METRICS={
    'input_tokens':('输入Token万',10000), 'output_tokens':('输出Token万',10000),
    'seconds':('耗时时分秒',86400), 'llm_calls':('LLM完成调用数次',1),
    'cache_rate':('缓存命中率%',1), 'tool_calls':('工具调用次',1),
    'tool_errors':('工具错误次',1)}
LABELS=[('平均输入Token（万）','input_tokens','token'),('平均输出Token（万）','output_tokens','token'),
    ('平均总Token（万）','total_tokens','token'),('平均端到端耗时','seconds','time'),
    ('端到端耗时P50','p50','time'),('端到端耗时P95','p95','time'),
    ('平均LLM完成调用数（次）','llm_calls','count'),('平均缓存命中率','cache_rate','rate'),
    ('平均工具调用数（次）','tool_calls','count'),('平均工具错误次数（次）','tool_errors','count')]


class InputError(ValueError): pass


def normalized(value):
    return re.sub(r'\s+','',str(value or '')).replace('（','(').replace('）',')')


def source_name(value):
    text=str(value or '').strip();key=text.casefold()
    if key in ('综合能力','常规能力'): return '综合能力'
    if key=='alphadata': return 'AlphaData'
    if key=='deepinsight': return 'DeepInsight'
    return text


def finite(value):
    return isinstance(value,(int,float)) and not isinstance(value,bool) and math.isfinite(value)


def metric_value(cell, key):
    value=cell.get('value');factor=METRICS[key][1]
    if value in (None,'','未取得','未记录','n.a.'): return None
    if key=='seconds' and isinstance(value,str):
        match=re.fullmatch(r'(?:(\d+)时)?(?:(\d+)分)?(\d+(?:\.\d+)?)秒',value)
        if match:return int(match[1] or 0)*3600+int(match[2] or 0)*60+float(match[3])
    if key=='cache_rate' and isinstance(value,str) and value.endswith('%'):
        try:value=float(value[:-1])/100
        except ValueError:raise InputError('无效百分比：'+value)
    if not finite(value) or value<0 or (key=='cache_rate' and value>1):
        raise InputError('指标包含无效数值：'+key+' '+str(value))
    return value*factor


def read_xlsx(path):
    """Resolve names/relationships, shared strings and cached formula values."""
    with ZipFile(path) as z:
        workbook=ET.fromstring(z.read('xl/workbook.xml'))
        props=workbook.find('s:workbookPr',NS)
        epoch=dt.datetime(1904,1,1) if props is not None and props.get('date1904') in ('1','true') else dt.datetime(1899,12,30)
        relations={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            shared=[''.join(x.itertext()) for x in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',NS)]
        sheets={}
        for sheet in workbook.findall('s:sheets/s:sheet',NS):
            target=relations[sheet.get('{'+R+'}id')].replace('\\','/')
            target=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            rows=[]
            for row in ET.fromstring(z.read(target)).findall('s:sheetData/s:row',NS):
                values={}
                for cell in row.findall('s:c',NS):
                    typ=cell.get('t');v=cell.find('s:v',NS)
                    if typ=='inlineStr':value=''.join(cell.find('s:is',NS).itertext())
                    elif typ=='s':value=shared[int(v.text)]
                    elif v is None or v.text is None:value=None
                    elif typ in ('str','e'):value=v.text
                    else:
                        try:value=float(v.text)
                        except ValueError:value=v.text
                    values[re.sub(r'\d','',cell.get('r'))]=dict(value=value,formula=cell.find('s:f',NS) is not None,
                        evidence=f'{path}#{sheet.get("name")}!{cell.get("r")}')
                rows.append(values)
            sheets[sheet.get('name')]=rows
    return sheets,epoch


def header(rows):
    for i,row in enumerate(rows):
        mapping={normalized(c['value']):column for column,c in row.items()}
        if {'题目来源','题号'}<=mapping.keys():return i,mapping
    raise InputError('工作表缺少题目来源和题号表头')


def value(row,columns,name):
    return row.get(columns.get(name),{}).get('value')


def key_of(row,columns):
    source=source_name(value(row,columns,'题目来源'));number=str(value(row,columns,'题号') or '').strip()
    if not source or not number:return None
    if source=='未识别' or not re.fullmatch(r'[1-9]\d*-[1-9]\d*',number):
        raise InputError('需要标准题号及题目来源：'+source+' '+number)
    return source,number


def get_metric(row,columns,key,issues):
    label=METRICS[key][0]
    if label not in columns:raise InputError('缺少指标列：'+label)
    cell=row.get(columns[label],{})
    if cell.get('formula') and cell.get('value') is None:
        raise InputError('公式缺少缓存值，请在Office中重算工作簿副本：'+cell['evidence'])
    result=metric_value(cell,key)
    if result is None:issues.append(dict(metric=key,evidence=cell.get('evidence'),reason='缺失数值'))
    return result


def load_workbook(path,name=None):
    path=Path(path).resolve();sheets,epoch=read_xlsx(path);summary=sheets.get('任务均值',next(iter(sheets.values())))
    hi,columns=header(summary);tasks={};issues=[]
    for row in summary[hi+1:]:
        key=key_of(row,columns)
        if key is None:continue
        if key in tasks:raise InputError('同一工作簿出现重复来源和题号：'+str(key))
        title=str(value(row,columns,'题目') or '').strip()
        if not title:raise InputError('题目名称为空：'+str(key))
        metrics={k:get_metric(row,columns,k,issues) for k in METRICS}
        attempts=value(row,columns,'独立次数')
        if not finite(attempts) or attempts<1 or not float(attempts).is_integer():raise InputError('独立次数无效：'+str(key))
        tasks[key]=dict(title=title,attempts=int(attempts),metrics=metrics)
    if not tasks:raise InputError('未找到可用题目：'+str(path))
    raw=[]
    if '逐次记录' in sheets:
        ri,rc=header(sheets['逐次记录']);ids=set()
        for row in sheets['逐次记录'][ri+1:]:
            key=key_of(row,rc)
            if key is None:continue
            rid=value(row,rc,'测试ID')
            if not rid or rid in ids:raise InputError('逐次记录缺少或重复测试ID：'+str(rid))
            ids.add(rid);metrics={k:get_metric(row,rc,k,issues) for k in METRICS}
            begin=value(row,rc,'开始时间');date=None
            if finite(begin):date=(epoch+dt.timedelta(days=begin)).date().isoformat()
            elif isinstance(begin,str) and re.match(r'^\d{4}-\d{2}-\d{2}',begin):date=begin[:10]
            raw.append(dict(key=key,run_id=str(rid),metrics=metrics,date=date))
    if raw:
        by_task=collections.defaultdict(list)
        for row in raw:by_task[row['key']].append(row)
        for k,t in tasks.items():
            if len(by_task[k])!=t['attempts']:raise InputError('任务均值与逐次记录的执行次数不一致：'+str(k))
            for metric in METRICS:
                values=[r['metrics'][metric] for r in by_task[k] if r['metrics'][metric] is not None]
                expected=statistics.mean(values) if values else None;actual=t['metrics'][metric]
                if actual is not None and (expected is None or not math.isclose(actual,expected,rel_tol=1e-8,abs_tol=1e-6)):
                    raise InputError('任务均值与逐次记录不一致：'+str(k)+' '+metric)
    if name is None:
        title=next((str(c['value']) for row in summary[:hi] for c in row.values() if '端侧日志统计' in str(c['value'])),None)
        name=title.replace('端侧日志统计','').strip() if title else path.stem.replace('_过程指标统计','')
    return dict(name=name,path=str(path),sha256=hashlib.sha256(path.read_bytes()).hexdigest(),tasks=tasks,raw=raw,issues=issues)


def percentile(values,p):
    if not values:return None
    ordered=sorted(values);at=(len(ordered)-1)*p;lo=math.floor(at);hi=math.ceil(at)
    return ordered[lo]+(ordered[hi]-ordered[lo])*(at-lo)


def compare(workbooks,focus='MobileWork'):
    if len(workbooks)<2:raise InputError('至少需要两个受测对象的工作簿')
    names=[w['name'] for w in workbooks]
    if len(set(names))!=len(names):raise InputError('同名对象存在多个文件，请明确选择批次或用--input 对象名=路径指定名称')
    if len({w['sha256'] for w in workbooks})!=len(workbooks):raise InputError('同一工作簿被重复提供')
    common=set.intersection(*(set(w['tasks']) for w in workbooks))
    if not common:raise InputError('各对象没有共同题目')
    for key in common:
        if len({w['tasks'][key]['title'] for w in workbooks})!=1:raise InputError('同号题目名称冲突：'+str(key))
    eligible={k:sorted(t for t in common if all(w['tasks'][t]['metrics'][k] is not None for w in workbooks)) for k in METRICS}
    eligible['total_tokens']=sorted(t for t in common if all(w['tasks'][t]['metrics'][k] is not None for w in workbooks for k in ('input_tokens','output_tokens')))
    result=dict(schema_version='1.0',focus=focus if focus in names else names[0],common_tasks=[dict(source=k[0],number=k[1],title=workbooks[0]['tasks'][k]['title']) for k in sorted(common)],objects=[],
        metric_task_counts={k:len(v) for k,v in eligible.items()})
    for w in workbooks:
        means={k:statistics.mean(w['tasks'][t]['metrics'][k] for t in eligible[k]) if eligible[k] else None for k in METRICS}
        means['total_tokens']=statistics.mean(sum(w['tasks'][t]['metrics'][k] for k in ('input_tokens','output_tokens')) for t in eligible['total_tokens']) if eligible['total_tokens'] else None
        raw=[r for r in w['raw'] if r['key'] in common];times=[r['metrics']['seconds'] for r in raw if r['metrics']['seconds'] is not None]
        means.update(p50=percentile(times,.5),p95=percentile(times,.95))
        def ratio(numerator,denominator):
            rows=[r['metrics'] for r in raw if r['metrics'][numerator] is not None and r['metrics'][denominator] is not None]
            den=sum(r[denominator] for r in rows)
            return sum(r[numerator] for r in rows)/den if den else None
        result['objects'].append(dict(name=w['name'],means=means,run_count=sum(w['tasks'][k]['attempts'] for k in common),time_n=len(times),
            input_per_call=ratio('input_tokens','llm_calls'),tool_error_rate=ratio('tool_errors','tool_calls'),
            dates=sorted({r['date'] for r in raw if r['date']}),
            tasks=[dict(source=k[0],number=k[1],**w['tasks'][k]) for k in sorted(common)],
            excluded_tasks=[list(k) for k in sorted(set(w['tasks'])-common)],issues=w['issues'],source=w['path'],source_sha256=w['sha256']))
    return result


def parser():
    p=argparse.ArgumentParser(description='汇总成本统计Excel；源工作簿只读')
    p.add_argument('--input-dir');p.add_argument('--input',action='append',default=[],help='文件路径或对象名=文件路径，可重复')
    p.add_argument('--focus',default='MobileWork');p.add_argument('--output',required=True)
    return p


def main():
    args=parser().parse_args();inputs=[]
    if args.input_dir:
        paths=sorted(p for p in Path(args.input_dir).glob('*.xlsx') if not p.name.startswith('~$'))
        priority={'MobileWork':0,'WorkBuddy':1,'TeleAgent':2}
        paths.sort(key=lambda p:(priority.get(p.stem.replace('_过程指标统计',''),3),p.name))
        inputs.extend((p,None) for p in paths if p.stem.endswith('_过程指标统计'))
    for item in args.input:
        if '=' in item:name,path=item.split('=',1);inputs.append((Path(path),name))
        else:inputs.append((Path(item),None))
    if not inputs:raise InputError('未找到输入Excel，请指定--input-dir或--input')
    data=compare([load_workbook(path,name) for path,name in inputs],args.focus)
    target=Path(args.output);target.parent.mkdir(parents=True,exist_ok=True)
    target.write_text(json.dumps(data,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf8')
    print(f'{len(data["objects"])}个对象，{len(data["common_tasks"])}道共同题目：{target.resolve()}')


if __name__=='__main__':
    try:main()
    except (InputError,OSError,KeyError,ET.ParseError) as exc:raise SystemExit('汇总停止：'+str(exc))
