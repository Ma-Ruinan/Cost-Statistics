"""Read accounting workbooks without Office, modifying neither files nor formulas."""
import argparse
import collections
import copy
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


def object_name(value):
    text=str(value or '').strip()
    return {'mobilework':'MobileWork','workbuddy':'WorkBuddy','teleagent':'TeleAgent'}.get(text.casefold(),text)


def timestamp(value,epoch):
    if value in (None,''):return None
    if finite(value):return (epoch+dt.timedelta(days=value)).isoformat(timespec='milliseconds')
    try:return dt.datetime.fromisoformat(str(value).replace('Z','+00:00')).isoformat(timespec='milliseconds')
    except ValueError:raise InputError('无法解析运行时间：'+str(value))


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
            start=timestamp(value(row,rc,'开始时间'),epoch);end=timestamp(value(row,rc,'结束时间'),epoch)
            raw.append(dict(key=key,run_id=str(rid),metrics=metrics,date=start[:10] if start else None,
                start=start,end=end,session_id=str(value(row,rc,'会话ID') or '').strip(),
                user_message_id=str(value(row,rc,'用户消息ID') or '').strip()))
    if '逐次记录' in sheets:
        by_task=collections.defaultdict(list)
        for row in raw:by_task[row['key']].append(row)
        if set(by_task)-tasks.keys():raise InputError('逐次记录中存在任务均值表未列出的题目：'+str(sorted(set(by_task)-tasks.keys())))
        for k,t in tasks.items():
            if len(by_task[k])!=t['attempts']:raise InputError('任务均值与逐次记录的执行次数不一致：'+str(k))
            for metric in METRICS:
                values=[r['metrics'][metric] for r in by_task[k] if r['metrics'][metric] is not None]
                expected=statistics.mean(values) if values else None;actual=t['metrics'][metric]
                if actual is not None and (expected is None or not math.isclose(actual,expected,rel_tol=1e-8,abs_tol=1e-6)):
                    raise InputError('任务均值与逐次记录不一致：'+str(k)+' '+metric)
    if name is None:
        title=next((str(c['value']) for row in summary[:hi] for c in row.values() if '端侧日志统计' in str(c['value'])),None)
        name=title.replace('端侧日志统计','').strip() if title else path.stem.split('_过程指标统计')[0]
    name=object_name(name)
    if not name:raise InputError('无法识别受测对象，请用--input 对象名=路径指定：'+str(path))
    return dict(name=name,path=str(path),sha256=hashlib.sha256(path.read_bytes()).hexdigest(),tasks=tasks,raw=raw,issues=issues)


def merge_workbooks(workbooks):
    """Union tasks per object before intersection; identify runs by native message IDs."""
    groups=collections.defaultdict(list);hashes={}
    for w in workbooks:
        name=object_name(w['name']);old=hashes.setdefault(w['sha256'],name)
        if old!=name:raise InputError('同一工作簿被指定为不同对象：'+old+'、'+name)
        groups[name].append(w)
    merged=[]
    priority={'MobileWork':0,'WorkBuddy':1,'TeleAgent':2}
    for name,items in sorted(groups.items(),key=lambda item:(priority.get(item[0],3),item[0])):
        unique=[];seen_files={};duplicates=[]
        for w in items:
            if w['sha256'] in seen_files:
                duplicates.append(dict(path=w['path'],duplicate_of=seen_files[w['sha256']]))
            else:seen_files[w['sha256']]=w['path'];unique.append(w)
        tasks=collections.defaultdict(list);raw_by_file={};runs=[];identities={};duplicate_runs=0
        for w in unique:
            by_task=collections.defaultdict(list);file_ids=set()
            for r in w['raw']:
                row=copy.deepcopy(r);row['evidence']=[dict(file=w['path'],run_id=row['run_id'])]
                stable=(row.get('session_id'),row.get('user_message_id'))
                if all(stable):
                    if stable in file_ids:raise InputError('同一工作簿重复列出同一次执行：'+str(stable)+' '+w['path'])
                    file_ids.add(stable)
                    if stable in identities:
                        previous=identities[stable]
                        if previous['key']!=row['key']:
                            raise InputError('同一执行被关联到不同题目：'+str(stable))
                        for field in ('start','end'):
                            a,b=previous.get(field),row.get(field)
                            if a is not None and b is not None and a!=b:
                                raise InputError('同一执行的时间冲突：'+str(stable)+' '+field+' '+w['path'])
                            if a is None:previous[field]=b
                        for metric in METRICS:
                            a,b=previous['metrics'][metric],row['metrics'][metric]
                            if a is not None and b is not None and not math.isclose(a,b,rel_tol=1e-8,abs_tol=1e-6):
                                raise InputError('同一执行的指标冲突：'+str(stable)+' '+metric+' '+w['path'])
                            if a is None:previous['metrics'][metric]=b
                        previous['date']=previous.get('date') or row.get('date')
                        previous['evidence'].extend(row['evidence']);duplicate_runs+=1
                    else:identities[stable]=row;runs.append(row)
                else:runs.append(row)
                by_task[row['key']].append(row)
            raw_by_file[w['path']]=by_task
            for key,task in w['tasks'].items():tasks[key].append((w,task))
        by_task=collections.defaultdict(list)
        for r in runs:by_task[r['key']].append(r)
        combined={}
        for key,parts in tasks.items():
            if len({t['title'] for _,t in parts})!=1:
                raise InputError('同一对象同号题目名称冲突：'+name+' '+str(key))
            if len(parts)==1:
                combined[key]=copy.deepcopy(parts[0][1]);continue
            for w,task in parts:
                rows=raw_by_file[w['path']][key]
                if len(rows)!=task['attempts']:
                    raise InputError('合并同题多份结果需要完整逐次记录：'+name+' '+str(key)+' '+w['path'])
                if any(not r.get('session_id') or not r.get('user_message_id') for r in rows):
                    raise InputError('合并同题多份结果需要会话ID和用户消息ID，测试ID不能用来去重：'+name+' '+str(key)+' '+w['path'])
            metrics={}
            for metric in METRICS:
                values=[r['metrics'][metric] for r in by_task[key] if r['metrics'][metric] is not None]
                metrics[metric]=statistics.mean(values) if values else None
            combined[key]=dict(title=parts[0][1]['title'],attempts=len(by_task[key]),metrics=metrics)
        merged.append(dict(name=name,tasks=combined,raw=runs,
            sources=[dict(path=w['path'],sha256=w['sha256']) for w in unique],
            issues=[i for w in unique for i in w['issues']],
            merge=dict(input_files=len(items),unique_files=len(unique),duplicate_files=duplicates,deduplicated_run_rows=duplicate_runs)))
    return merged


def percentile(values,p):
    if not values:return None
    ordered=sorted(values);at=(len(ordered)-1)*p;lo=math.floor(at);hi=math.ceil(at)
    return ordered[lo]+(ordered[hi]-ordered[lo])*(at-lo)


def compare(workbooks,focus='MobileWork'):
    workbooks=merge_workbooks(workbooks)
    if len(workbooks)<2:raise InputError('至少需要两个不同受测对象的工作簿')
    names=[w['name'] for w in workbooks]
    common=set.intersection(*(set(w['tasks']) for w in workbooks))
    if not common:raise InputError('各对象没有共同题目')
    for key in common:
        if len({w['tasks'][key]['title'] for w in workbooks})!=1:raise InputError('同号题目名称冲突：'+str(key))
    eligible={k:sorted(t for t in common if all(w['tasks'][t]['metrics'][k] is not None for w in workbooks)) for k in METRICS}
    eligible['total_tokens']=sorted(t for t in common if all(w['tasks'][t]['metrics'][k] is not None for w in workbooks for k in ('input_tokens','output_tokens')))
    focus=object_name(focus)
    result=dict(schema_version='1.1',focus=focus if focus in names else names[0],common_tasks=[dict(source=k[0],number=k[1],title=workbooks[0]['tasks'][k]['title']) for k in sorted(common)],objects=[],
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
            excluded_tasks=[list(k) for k in sorted(set(w['tasks'])-common)],issues=w['issues'],sources=w['sources'],merge=w['merge'],
            runs=[dict(source=r['key'][0],number=r['key'][1],**{k:v for k,v in r.items() if k!='key'}) for r in raw]))
    return result


def parser():
    p=argparse.ArgumentParser(description='汇总成本统计Excel；源工作簿只读')
    p.add_argument('--input-dir',help='递归发现过程指标统计Excel');p.add_argument('--input',action='append',default=[],help='文件路径或对象名=文件路径，可重复；同对象自动合并')
    p.add_argument('--focus',default='MobileWork');p.add_argument('--output',required=True)
    return p


def main():
    args=parser().parse_args();inputs=[]
    if args.input_dir:
        paths=sorted(p for p in Path(args.input_dir).rglob('*.xlsx') if not p.name.startswith('~$'))
        inputs.extend((p,None) for p in paths if '_过程指标统计' in p.stem)
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
