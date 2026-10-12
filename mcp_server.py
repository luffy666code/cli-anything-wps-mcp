"""Compact stdio MCP server, Windows Python stdlib only (pywin32 needed only for export).

MCP JSON-RPC protocol (2024-11-05 / 2025-03-26 subset): initialize, ping,
tools/list and tools/call. Each request is one line of UTF-8 JSON. No stdout logs.
"""
from __future__ import annotations
import json
import re
import sys
import traceback
from collections.abc import Callable
from typing import Any
import project_store as db
import com_export

TOOLS: dict[str, tuple[dict, Callable]] = {}

def f(name: str, desc: str, props: dict, required: list[str] | None = None):
    def decorate(fn: Callable):
        TOOLS[name] = ({'name': name, 'description': desc, 'inputSchema': {
            'type': 'object', 'properties': props, 'required': required or [], 'additionalProperties': False}}, fn)
        return fn
    return decorate

def string(desc: str, **kw) -> dict: return {'type': 'string', 'description': desc, **kw}
def number(desc: str, **kw) -> dict: return {'type': 'number', 'description': desc, **kw}
def integer(desc: str, **kw) -> dict: return {'type': 'integer', 'description': desc, **kw}
def boolean(desc: str) -> dict: return {'type': 'boolean', 'description': desc}
def arr(desc: str) -> dict: return {'type': 'array', 'description': desc}
ID = {'project_id': string('project_create 返回的项目 ID（32 位十六进制）')}


def obj_of(args: dict, required_type: str | None = None) -> dict:
    obj = db.load(args['project_id'])
    if required_type and obj['type'] != required_type:
        raise ValueError(f'此操作需要 {required_type} 项目，当前为 {obj["type"]}')
    return obj


def stored(obj: dict) -> dict:
    db.save(obj)
    return db.summary(obj)


@f('wps_cli_project_create', '创建一个持久化 JSON 文档项目，返回 project_id；writer=Word，calc=Excel，impress=PPT。',
   {'doc_type': string('文档类型', enum=['writer', 'calc', 'impress']),
    'name': string('项目名称'), 'slide_width': number('仅 PPT：幻灯片宽度，单位 pt；默认 960'),
    'slide_height': number('仅 PPT：幻灯片高度，单位 pt；默认 540')}, ['doc_type', 'name'])
def project_create(a):
    w, h = float(a.get('slide_width', 960)), float(a.get('slide_height', 540))
    if not 100 <= w <= 2500 or not 100 <= h <= 2500: raise ValueError('幻灯片宽高必须在 100~2500pt 内')
    return db.create(a['doc_type'], a['name'], w, h)


@f('wps_cli_project_list', '列出当前用户工作目录中的项目。', {})
def project_list(a): return db.all_projects()


@f('wps_cli_project_info', '查看项目摘要或完整 JSON 文档数据。', {**ID, 'full': boolean('是否返回完整项目模型')}, ['project_id'])
def project_info(a):
    obj = obj_of(a)
    if a.get('full'): return {k:v for k,v in obj.items() if k not in ('undo_stack', 'redo_stack')}
    return db.summary(obj)


@f('wps_cli_project_delete', '永久删除一个 JSON 项目。此操作不可撤销，需要 confirm=true。',
   {**ID,'confirm':boolean('确认永久删除该项目，必须为 true')}, ['project_id','confirm'])
def project_delete(a):
    if a['confirm'] is not True: raise ValueError('永久删除项目需要 confirm=true')
    return db.remove(a['project_id'])


@f('wps_cli_undo', '撤销最近一次项目内容操作（保留最近 30 个快照）。', ID, ['project_id'])
def undo(a): return db.undo_redo(a['project_id'], True)


@f('wps_cli_redo', '重做最近一次已撤销的项目内容操作。', ID, ['project_id'])
def redo(a): return db.undo_redo(a['project_id'], False)


@f('wps_cli_writer_add_text', '向 Word JSON 项目追加段落或标题，导出 DOCX/PDF 时再写入 WPS。',
   {**ID, 'text': string('文本'), 'kind': string('类型', enum=['paragraph','heading']),
    'level': integer('标题级别 1~6'), 'font_size': number('字号 pt'), 'font_name': string('字体'),
    'color': string('颜色 #RRGGBB'), 'bold': boolean('是否加粗'),
    'alignment': string('对齐', enum=['left','center','right','justify'])}, ['project_id','text'])
def writer_add_text(a):
    o = obj_of(a,'writer')
    t = a['text']
    if len(t)>100000: raise ValueError('单次文本不得超过 10 万字')
    kind = a.get('kind','paragraph')
    if kind not in ('paragraph','heading'): raise ValueError('kind 必须 paragraph 或 heading')
    level = int(a.get('level',1))
    if not 1<=level<=6: raise ValueError('标题级别应为1~6')
    db.snapshot(o)
    item = {'kind':kind,'text':t,'level':level}
    for k in ('font_size','font_name','color','bold','alignment'):
        if k in a: item[k]=a[k]
    o['content'].append(item)
    return {**stored(o),'appended':item}


@f('wps_cli_writer_add_table', '向 Word 项目添加指定行列的表格，可选填充二维数据。',
   {**ID,'rows':integer('行数'), 'cols':integer('列数'),
    'data':arr('二维数组，每行一个数组，长度不得超过 cols')}, ['project_id','rows','cols'])
def writer_add_table(a):
    o = obj_of(a,'writer')
    r,c=int(a['rows']),int(a['cols'])
    if not 1<=r<=300 or not 1<=c<=50: raise ValueError('行数1~300，列数1~50')
    data=a.get('data',[])
    if not isinstance(data,list) or any(not isinstance(row,list) or len(row)>c for row in data) or len(data)>r:
        raise ValueError('data 必须是未超出表格行列范围的二维数组')
    db.snapshot(o);o['content'].append({'kind':'table','rows':r,'cols':c,'data':data})
    return stored(o)


@f('wps_cli_writer_add_image', '在 Word 项目中记录本机图片路径，导出时读取原图。',
   {**ID,'path':string('Windows 本机图片完整路径'), 'width':number('可选，宽度pt'),
    'height':number('可选，高度pt')}, ['project_id','path'])
def writer_add_image(a):
    o=obj_of(a,'writer');item={'kind':'image','path':a['path']}
    for k in ('width','height'):
        if k in a: item[k]=a[k]
    db.snapshot(o);o['content'].append(item)
    return stored(o)


@f('wps_cli_writer_page_break', '在 Word 项目中追加分页符。', ID, ['project_id'])
def writer_page_break(a):
    o=obj_of(a,'writer');db.snapshot(o);o['content'].append({'kind':'page_break'});return stored(o)


@f('wps_cli_writer_find_replace', '在项目文本模型中查找并替换，不修改已存在的 DOCX 文件。',
   {**ID,'find_text':string('待查找文字'),'replace_text':string('替换文字')}, ['project_id','find_text','replace_text'])
def writer_find_replace(a):
    o=obj_of(a,'writer');src=a['find_text']
    if not src: raise ValueError('find_text 不能为空')
    db.snapshot(o);cnt=0
    for item in o['content']:
        if 'text' in item:
            cnt+=item['text'].count(src);item['text']=item['text'].replace(src,a['replace_text'])
    db.save(o)
    return {'replacements':cnt,**db.summary(o)}


@f('wps_cli_writer_list', '列出 Word 项目内容模型。', ID, ['project_id'])
def writer_list(a): return obj_of(a,'writer')['content']


@f('wps_cli_calc_set_cell', '在 Excel JSON 项目中设置单元格的值或公式。',
   {**ID,'address':string('A1, B23 等单元格地址'),
    'value':{'description':'字符串、数字、布尔或 null；设置公式时可省略'},
    'formula':string('可选 Excel 公式，如 =SUM(A1:A5)'),
    'sheet':integer('工作表索引，0 开始，默认0')}, ['project_id','address'])
def calc_set_cell(a):
    o=obj_of(a,'calc');ref=a['address'].upper()
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*',ref):raise ValueError('单元格地址必须是 A1 或 AA10 格式')
    index=int(a.get('sheet',0))
    if index<0 or index>=len(o['sheets']): raise ValueError('sheet 工作表索引越界')
    db.snapshot(o)
    cell={'value':a.get('value')}
    if a.get('formula'):cell['formula']=a['formula']
    o['sheets'][index]['cells'][ref]=cell
    return {**stored(o),'cell':{'address':ref,**cell}}


@f('wps_cli_calc_set_range', '将二维数组填入指定工作表的连续区域。',
   {**ID,'start_cell':string('起始单元格，例如 A1'),'data':arr('二维数组'),
    'sheet':integer('工作表索引，默认0')}, ['project_id','start_cell','data'])
def calc_set_range(a):
    o=obj_of(a,'calc');match=re.fullmatch(r'([A-Za-z]{1,3})([1-9][0-9]*)',a['start_cell'])
    if not match:raise ValueError('start_cell 格式应为 A1')
    data=a['data']
    if not isinstance(data,list) or not data or len(data)>1000 or any(not isinstance(r,list) or len(r)>100 for r in data):
        raise ValueError('data 要求非空二维数组，最多1000行/100列')
    index=int(a.get('sheet',0))
    if index<0 or index>=len(o['sheets']):raise ValueError('sheet 工作表索引越界')
    first_col=0
    for ch in match.group(1).upper(): first_col=first_col*26+ord(ch)-64
    first_row=int(match.group(2));db.snapshot(o);cells=o['sheets'][index]['cells']
    count=0
    for ri,row in enumerate(data):
        for ci,val in enumerate(row):
            n=first_col+ci;letters=''
            while n: n,r=divmod(n-1,26);letters=chr(65+r)+letters
            cells[f'{letters}{first_row+ri}']={'value':val};count+=1
    return {**stored(o),'cells_written':count}


@f('wps_cli_calc_add_sheet', '给 Excel 项目追加工作表。',
   {**ID,'name':string('工作表名称')}, ['project_id','name'])
def calc_add_sheet(a):
    o=obj_of(a,'calc');name=a['name']
    if not name or len(name)>31 or any(ch in name for ch in r'[]:*?/\\'):
        raise ValueError('工作表名不能为空、不能超过31字符或包含 []:*?/\\')
    if any(s['name'].casefold()==name.casefold() for s in o['sheets']):raise ValueError('工作表名称重复')
    db.snapshot(o);o['sheets'].append({'name':name,'cells':{},'merges':[]});return stored(o)


@f('wps_cli_calc_merge', '标记 Excel 单元格区域合并，导出时应用。',
   {**ID,'range':string('范围如 A1:C1'),'sheet':integer('工作表索引，默认0')}, ['project_id','range'])
def calc_merge(a):
    o=obj_of(a,'calc');rg=a['range'].upper()
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*',rg):raise ValueError('range 格式应为 A1:C1')
    i=int(a.get('sheet',0))
    if i<0 or i>=len(o['sheets']):raise ValueError('工作表索引越界')
    db.snapshot(o);o['sheets'][i]['merges'].append({'range':rg});return stored(o)


@f('wps_cli_calc_get_cell', '从 JSON 项目模型读取单元格的值或公式。',
   {**ID,'address':string('单元格地址'),'sheet':integer('工作表索引，默认0')}, ['project_id','address'])
def calc_get_cell(a):
    o=obj_of(a,'calc');i=int(a.get('sheet',0));ref=a['address'].upper()
    if i<0 or i>=len(o['sheets']):raise ValueError('sheet 工作表索引越界')
    return {'address':ref,'cell':o['sheets'][i]['cells'].get(ref)}


@f('wps_cli_calc_sheets', '列出 Excel JSON 项目的工作表及非空单元格数。',ID,['project_id'])
def calc_sheets(a):
    o=obj_of(a,'calc')
    return [{'index':i,'name':s['name'],'cell_count':len(s['cells'])} for i,s in enumerate(o['sheets'])]


@f('wps_cli_ppt_add_slide', '新增一张 PPT 幻灯片，后续可追加文字、图片或形状。',
   {**ID,'title':string('标题'),'content':string('正文'),'background':string('可选背景色 #RRGGBB')}, ['project_id'])
def ppt_add_slide(a):
    o=obj_of(a,'impress');db.snapshot(o)
    s={'title':a.get('title',''),'content':a.get('content',''),'elements':[]}
    if a.get('background'):s['background']=a['background']
    o['slides'].append(s)
    return {**stored(o),'slide_index':len(o['slides'])-1}


def slide_ref(a: dict) -> tuple[dict,dict]:
    o=obj_of(a,'impress');i=int(a['slide_index'])
    if i<0 or i>=len(o['slides']):raise ValueError('slide_index 超出范围')
    return o,o['slides'][i]


def coords(a: dict) -> dict:
    d={k:float(a[k]) for k in ('x','y','width','height')}
    if any(abs(v)>10000 for v in d.values()) or d['width']<=0 or d['height']<=0:
        raise ValueError('元素尺寸或位置不合法')
    return d


@f('wps_cli_ppt_add_text', '向已有幻灯片增加真正会参与 PPTX 渲染的文本框。',
   {**ID,'slide_index':integer('幻灯片索引，从0开始'),
    'text':string('文字'),'x':number('左侧位置 pt'),'y':number('顶部位置 pt'),
    'width':number('宽度 pt'),'height':number('高度 pt'),
    'font_size':number('字号 pt'),'color':string('颜色 #RRGGBB'),'bold':boolean('是否加粗')},
   ['project_id','slide_index','text','x','y','width','height'])
def ppt_add_text(a):
    o,s=slide_ref(a);e={'kind':'text','text':a['text'],**coords(a)}
    for k in ('font_size','color','bold'):
        if k in a:e[k]=a[k]
    db.snapshot(o);s['elements'].append(e)
    return {**stored(o),'element_index':len(s['elements'])-1}


@f('wps_cli_ppt_add_shape', '向 PPT 幻灯片添加矩形、圆角矩形或椭圆形状。',
   {**ID,'slide_index':integer('幻灯片索引'),'shape':string('图形',enum=['rect','roundrect','ellipse']),
    'x':number('X pt'),'y':number('Y pt'),'width':number('宽 pt'),'height':number('高 pt'),
    'fill':string('填充色 #RRGGBB'),'text':string('图形内文字')},
   ['project_id','slide_index','shape','x','y','width','height'])
def ppt_add_shape(a):
    o,s=slide_ref(a);e={'kind':'shape','shape':a['shape'],**coords(a)}
    for k in ('fill','text'):
        if k in a:e[k]=a[k]
    db.snapshot(o);s['elements'].append(e)
    return {**stored(o),'element_index':len(s['elements'])-1}


@f('wps_cli_ppt_add_image', '向 PPT 幻灯片加入本机图片路径，导出时读取。',
   {**ID,'slide_index':integer('幻灯片索引'),'path':string('图片绝对路径'),
    'x':number('X pt'),'y':number('Y pt'),'width':number('宽 pt'),'height':number('高 pt')},
   ['project_id','slide_index','path','x','y','width','height'])
def ppt_add_image(a):
    o,s=slide_ref(a);db.snapshot(o)
    s['elements'].append({'kind':'image','path':a['path'],**coords(a)})
    return {**stored(o),'element_index':len(s['elements'])-1}


@f('wps_cli_ppt_list', '查看 PPT JSON 项目中的幻灯片标题及元素数量。', ID, ['project_id'])
def ppt_list(a):
    o=obj_of(a,'impress')
    return [{'index':i,'title':s.get('title'),'element_count':len(s.get('elements',[]))}
            for i,s in enumerate(o['slides'])]


@f('wps_cli_export', '使用 Windows 本机 WPS COM 将 JSON 模型导出为 DOCX、XLSX、PPTX 或 PDF。',
   {**ID,'output_path':string('完整 Windows 文件路径，扩展名必须与 format 一致'),
    'format':string('输出格式',enum=['docx','xlsx','pptx','pdf']),
    'overwrite':boolean('文件已存在时是否覆盖，默认false')},
   ['project_id','output_path','format'])
def export_document(a):
    obj=obj_of(a)
    return com_export.export_project(obj, a['output_path'],a['format'],bool(a.get('overwrite',False)))


def _send(obj: dict) -> None:
    # stdio JSON-RPC: never print anything else to stdout.
    sys.stdout.write(json.dumps(obj, ensure_ascii=False, separators=(',',':'))+'\n')
    sys.stdout.flush()


def _message(raw: dict) -> None:
    mid = raw.get('id')
    method = raw.get('method','')
    if 'id' not in raw:
        return # Notification, including notifications/initialized.
    try:
        if method=='initialize':
            version=raw.get('params',{}).get('protocolVersion','2024-11-05')
            if version not in ('2024-11-05','2025-03-26','2025-06-18','2025-11-25'):
                version='2024-11-05'
            result={'protocolVersion':version,
                    'capabilities':{'tools':{'listChanged':False}},
                    'serverInfo':{'name':'cli-anything-wps','version':'1.0.0'},
                    'instructions':'Windows-only WPS COM document generator. Start with wps_cli_project_create, edit by project_id, finish with wps_cli_export. Editing changes JSON, not live open files.'}
        elif method=='ping': result={}
        elif method=='tools/list': result={'tools':[item[0] for item in TOOLS.values()]}
        elif method=='tools/call':
            p=raw.get('params',{})
            name=p.get('name','')
            if name not in TOOLS: raise ValueError('未知工具: '+name)
            args=p.get('arguments') or {}
            if not isinstance(args,dict):raise ValueError('arguments 必须为 JSON 对象')
            schema,fn=TOOLS[name]
            missing=[k for k in schema['inputSchema']['required'] if k not in args]
            if missing: raise ValueError('缺少必需参数: '+', '.join(missing))
            bad=set(args)-set(schema['inputSchema']['properties'])
            if bad: raise ValueError('不支持的参数: '+', '.join(sorted(bad)))
            try:
                result_value=fn(args)
                result={'content':[{'type':'text','text':json.dumps(result_value,ensure_ascii=False,default=str)}],
                        'isError':False}
            except Exception as ex:
                print(f'[WPS-CLI-MCP] {name}: {type(ex).__name__}: {ex}', file=sys.stderr)
                result={'content':[{'type':'text','text':f'{type(ex).__name__}: {ex}'}],'isError':True}
        else:
            _send({'jsonrpc':'2.0','id':mid,'error':{'code':-32601,'message':'Method not found: '+method}})
            return
        _send({'jsonrpc':'2.0','id':mid,'result':result})
    except Exception as ex:
        _send({'jsonrpc':'2.0','id':mid,'error':{'code':-32603,'message':str(ex)}})


def run():
    # Windows terminal encoding may be CP936; MCP stdio is always UTF-8.
    if hasattr(sys.stdin, 'reconfigure'):
        sys.stdin.reconfigure(encoding='utf-8')
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    for line in sys.stdin:
        if not line.strip():continue
        try:
            obj=json.loads(line)
            if not isinstance(obj,dict):raise ValueError('请求必须为JSON对象')
            _message(obj)
        except Exception as ex:
            print(f'[WPS-CLI-MCP] protocol error: {ex}',file=sys.stderr)


if __name__=='__main__': run()
