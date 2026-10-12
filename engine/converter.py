#!/usr/bin/env python3
"""Local document conversion. Only staging copies are opened by the engine."""
from __future__ import annotations
import argparse, hashlib, json, os, signal, struct, subprocess, tempfile, sys
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile, ZipInfo

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
OD = '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
TX = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
TB = '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}'
DR = '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}'
MAX_SIZE = 100 * 1024 * 1024
MAX_UNPACKED = 256 * 1024 * 1024
sys.path.insert(0, str(Path(__file__).resolve().parent))

class ConversionError(Exception): pass

def digest(data): return hashlib.sha256(data).hexdigest()
def u16(data, pos): return struct.unpack_from('<H', data, pos)[0]
def u32(data, pos): return struct.unpack_from('<I', data, pos)[0]

class CompoundFile:
    def __init__(self, data):
        self.data = data
        if len(data) < 512 or data[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
            raise ConversionError('文档容器损坏。')
        self.size = 1 << u16(data, 30)
        self.msize = 1 << u16(data, 32)
        if self.size not in (512, 4096) or self.msize != 64:
            raise ConversionError('不支持的旧文档容器。')
        self.sectors = max(0, len(data)//self.size-1)
        difat = list(struct.unpack_from('<109I', data, 76))
        sid = u32(data, 68)
        for _ in range(min(u32(data,72), self.sectors)):
            vals = struct.unpack('<%dI' % (self.size//4), self.sector(sid))
            difat.extend(vals[:-1]); sid = vals[-1]
        self.fat = []
        for sid in [v for v in difat if v < 0xfffffffa][:u32(data,44)]:
            self.fat.extend(struct.unpack('<%dI'%(self.size//4),self.sector(sid)))
        directory = self.regular(u32(data,48))
        self.entries = {}
        root = None
        for pos in range(0,len(directory),128):
            d = directory[pos:pos+128]
            if len(d) != 128 or not d[66]: continue
            n = u16(d,64)
            if not 2 <= n <= 64: raise ConversionError('旧文档目录损坏。')
            name = d[:n-2].decode('utf-16le')
            length = u32(d,120) if u16(data,26)==3 else struct.unpack_from('<Q',d,120)[0]
            entry = (u32(d,116),length)
            if d[66] == 2:
                if name in self.entries:
                    raise ConversionError('文档含嵌套对象，当前版本无法可靠区分其内容。')
                self.entries[name] = entry
            if d[66] == 5: root = entry
        if root is None: raise ConversionError('旧文档缺少根目录。')
        self.mini = self.regular(root[0])[:root[1]] if root[1] else b''
        mf = self.regular(u32(data,60)) if u32(data,64) else b''
        self.mfat = list(struct.unpack('<%dI'%(len(mf)//4),mf))

    def sector(self,sid):
        if not 0 <= sid < self.sectors: raise ConversionError('旧文档数据块越界。')
        return self.data[(sid+1)*self.size:(sid+2)*self.size]

    def chain(self,sid,table):
        seen = set()
        while sid not in (0xfffffffe,0xffffffff):
            if sid in seen or not 0 <= sid < len(table):
                raise ConversionError('旧文档数据链损坏。')
            seen.add(sid); yield sid; sid = table[sid]

    def regular(self,sid): return b''.join(self.sector(v) for v in self.chain(sid,self.fat))

    def stream(self,name):
        sid, length = self.entries[name]
        if length > MAX_UNPACKED: raise ConversionError('文档内部数据过大。')
        data = (b''.join(self.mini[v*self.msize:(v+1)*self.msize] for v in self.chain(sid,self.mfat))
                if length < u32(self.data,56) else self.regular(sid))
        if len(data) < length: raise ConversionError('文档内部数据不完整。')
        return data[:length]

def read_binary_word(data):
    cf = CompoundFile(data)
    if 'WordDocument' not in cf.entries:
        if 'EncryptedPackage' in cf.entries:
            raise ConversionError('这是加密文档，当前版本需要先取得未加密副本。')
        raise ConversionError('这不是支持的 Word 兼容文档；不能仅凭 WPS 后缀转换。')
    word = cf.stream('WordDocument')
    if len(word)<100 or u16(word,0)!=0xa5ec or u16(word,2)<0xc1:
        raise ConversionError('当前版本不支持这种更早期的 Word / WPS 格式。')
    flags=u16(word,10)
    if flags & 0x0100: raise ConversionError('这是加密文档，当前版本需要未加密副本。')
    if any('VBA' in k.upper() or k in ('Macros','PROJECT') for k in cf.entries):
        raise ConversionError('文档包含宏，转成 DOCX 会移除宏，当前版本保留原件并停止。')
    lwpos = 34+2*u16(word,32)
    n=u16(word,lwpos); start=lwpos+2
    if n<11 or n>64: raise ConversionError('无法可靠读取旧文档正文。')
    lw=list(struct.unpack_from('<%dI'%n,word,start))
    pairpos=start+4*n+2
    if u16(word,start+4*n)<=33: raise ConversionError('旧文档缺少正文索引。')
    fc,lcb=struct.unpack_from('<II',word,pairpos+33*8)
    table=cf.stream('1Table' if flags&0x200 else '0Table')
    clx=table[fc:fc+lcb]; at=0
    while at<len(clx) and clx[at]==1: at+=3+u16(clx,at+1)
    if at+5>len(clx) or clx[at]!=2: raise ConversionError('旧文档正文索引损坏。')
    length=u32(clx,at+1); plc=clx[at+5:at+5+length]
    if len(plc)!=length or length<4 or (length-4)%12: raise ConversionError('旧文档正文索引不完整。')
    count=(length-4)//12
    cps=struct.unpack_from('<%dI'%(count+1),plc,0)
    parts=[]
    for i in range(count):
        f=u32(plc,4*(count+1)+8*i+2); compressed=bool(f&0x40000000)
        offset=(f&0x3fffffff)//2 if compressed else f&0x3fffffff
        units=cps[i+1]-cps[i]
        if units<0 or units>MAX_UNPACKED: raise ConversionError('正文长度无效。')
        raw=word[offset:offset+units*(1 if compressed else 2)]
        if len(raw)!=units*(1 if compressed else 2): raise ConversionError('正文数据缺失。')
        # MS-DOC FcCompressed stores 8-bit Unicode plus the specified punctuation mapping.
        if compressed:
            mapping={b:ord(bytes([b]).decode('cp1252')) for b in
                     [0x82,0x83,0x84,0x85,0x86,0x87,0x88,0x89,0x8a,0x8b,0x8c,
                      0x91,0x92,0x93,0x94,0x95,0x96,0x97,0x98,0x99,0x9a,0x9b,0x9c,0x9f]}
            parts.append(''.join(chr(mapping.get(b,b)) for b in raw))
        else: parts.append(raw.decode('utf-16le'))
    # Slice stories by UTF-16 code units, including supplementary characters.
    encoded=''.join(parts).encode('utf-16le'); pos=0; stories={}
    for name,index in [('main',3),('footnote',4),('header',5),('annotation',7),
                       ('endnote',8),('textbox',9),('header_textbox',10)]:
        units=lw[index]; chunk=encoded[pos*2:(pos+units)*2]
        if len(chunk)!=units*2: raise ConversionError('文档正文不完整。')
        stories[name]=chunk.decode('utf-16le'); pos+=units
    return stories

def zip_members(data):
    import io
    z=ZipFile(io.BytesIO(data)); items=z.infolist()
    if sum(i.file_size for i in items)>MAX_UNPACKED: raise ConversionError('压缩文档内部数据过大。')
    names={}
    for info in items:
        name=info.filename.replace('\\','/')
        if name.startswith('/') or '..' in name.split('/') or name in names:
            raise ConversionError('压缩文档目录异常。')
        names[name]=info
    return z,names

def docx_info(data):
    z,names=zip_members(data)
    try:
        if 'word/document.xml' not in names or '[Content_Types].xml' not in names or '_rels/.rels' not in names:
            raise ConversionError('DOCX 缺少必要的内部文件。')
        if z.testzip() is not None: raise ConversionError('DOCX 内部校验失败，文件可能损坏。')
        for name in names:
            if name.endswith(('.xml','.rels')): ET.fromstring(z.read(names[name]))
        types=z.read(names['[Content_Types].xml'])
        if b'macroEnabled' in types or any('vbaProject' in k for k in names):
            raise ConversionError('文档包含宏，无法原样保存为 DOCX。')
        ct=ET.fromstring(types)
        if ct.tag!='{http://schemas.openxmlformats.org/package/2006/content-types}Types':
            raise ConversionError('DOCX 的内容类型目录无效。')
        main_types=[e.attrib.get('ContentType','') for e in ct
                    if e.attrib.get('PartName','').replace('\\','/')=='/word/document.xml']
        if not any('wordprocessingml.document.main+xml' in v for v in main_types):
            raise ConversionError('文档不是标准 DOCX 类型。')
        rel=ET.fromstring(z.read(names['_rels/.rels']))
        if rel.tag!='{http://schemas.openxmlformats.org/package/2006/relationships}Relationships':
            raise ConversionError('DOCX 的文档关系无效。')
        mainrels=[e for e in rel if e.attrib.get('Type','').endswith('/officeDocument')]
        if len(mainrels)!=1 or mainrels[0].attrib.get('Target','').replace('\\','/').lstrip('/')!='word/document.xml' or mainrels[0].attrib.get('TargetMode')=='External':
            raise ConversionError('DOCX 的主文档关系缺失或无效。')
        root=ET.fromstring(z.read(names['word/document.xml']))
        wp=root.tag.removesuffix('document')
        if wp not in (W,'{http://purl.oclc.org/ooxml/wordprocessingml/main}'):
            raise ConversionError('DOCX 的正文格式无效。')
        body=root.find(wp+'body')
        if body is None: raise ConversionError('DOCX 缺少正文。')
        def text(node):
            result=[]
            def walk(el):
                if el.tag in (wp+'pPr',wp+'txbxContent'): return
                if el.tag in (wp+'t',wp+'delText'): result.append(el.text or '')
                elif el.tag==wp+'tab': result.append('\t')
                elif el.tag==wp+'br': result.append('\n')
                else:
                    for child in el: walk(child)
            walk(node)
            return ''.join(result)
        paragraphs=[]
        def walk_body(el):
            if el.tag==wp+'txbxContent': return
            if el.tag==wp+'p': paragraphs.append(text(el)+'\r')
            else:
                for child in el: walk_body(child)
        walk_body(body)
        main=''.join(paragraphs)
        return {'main':main,'paragraphs':len(paragraphs),'tables':sum(1 for _ in body.iter(wp+'tbl')),
                'images':sum(1 for e in body.iter() if e.tag.endswith(('}blip','}imagedata'))),
                'image_hashes':sorted(digest(z.read(info)) for k,info in names.items() if k.startswith('word/media/') and not k.endswith('/')),
                'revisions':sum(1 for e in root.iter() if e.tag in (wp+'ins',wp+'del')),
                'comments':('word/comments.xml' in names),
                'fonts':sorted({v for e in root.iter(W+'rFonts') for k,v in e.attrib.items()
                                if k in (W+'ascii',W+'hAnsi',W+'eastAsia',W+'cs')}),
                'noncanonical_names':any(i.filename!=k for k,i in names.items())}
    finally: z.close()

def detect(data):
    if data[:8]==bytes.fromhex('d0cf11e0a1b11ae1'):
        return 'doc',read_binary_word(data)
    if data.lstrip().startswith(b'{\\rtf'): return 'rtf',None
    if data.startswith(b'PK'):
        z,names=zip_members(data)
        try:
            if 'word/document.xml' in names: return 'docx',docx_info(data)
            if 'mimetype' in names and z.read(names['mimetype']).strip()==b'application/vnd.oasis.opendocument.text':
                ET.fromstring(z.read(names['content.xml']))
                return 'odt',None
        finally: z.close()
        raise ConversionError('该压缩文件不是支持的文字文档。')
    raise ConversionError('无法识别为支持的 DOC、DOCX、RTF 或 ODT；已保留原件。')

def visible_binary(text):
    # Keep field results, omit field instructions and dynamic page values.
    out=[]; stack=[]
    for ch in text:
        if ch=='\x13': stack.append({'code':'','result':False,'dynamic':False}); continue
        if ch=='\x14' and stack:
            stack[-1]['result']=True
            stack[-1]['dynamic']=stack[-1]['code'].lstrip().upper().startswith(('PAGE ','PAGE\\','NUMPAGES '))
            continue
        if ch=='\x15' and stack: stack.pop(); continue
        if stack and not stack[-1]['result']: stack[-1]['code']+=ch; continue
        if any(s['dynamic'] for s in stack): continue
        if ch not in '\r\n\x07\x01\x02\x05\x08\x0b\x0c\x0e': out.append(ch)
    return ''.join(out)

def visible_xml(text): return text.replace('\r','').replace('\n','')

def odt_info(path):
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('content.xml')); body=root.find(OD+'body')
        if body is None: raise ConversionError('引擎未读取到正文。')
        texts=[]
        def flatten(el):
            s=el.text or ''
            for child in el:
                if child.tag in (TX+'note',OD+'annotation',DR+'text-box'): pass
                elif child.tag==TX+'s': s+=' '*int(child.attrib.get(TX+'c','1'))
                elif child.tag==TX+'tab': s+='\t'
                elif child.tag==TX+'line-break': s+='\n'
                else: s+=flatten(child)
                s+=child.tail or ''
            return s
        def walk(el):
            if el.tag in (TX+'note',OD+'annotation',DR+'text-box'): return
            if el.tag in (TX+'p',TX+'h'): texts.append(flatten(el)+'\r')
            else:
                for child in el: walk(child)
        walk(body)
        return {'main':''.join(texts),'tables':sum(1 for _ in body.iter(TB+'table')),
                'images':sum(1 for _ in body.iter(DR+'image')),
                'image_hashes':sorted(digest(z.read(k)) for k in z.namelist() if k.startswith('Pictures/') and not k.endswith('/')),
                'objects':sum(1 for e in body.iter() if e.tag in (DR+'object',DR+'object-ole'))}

def convert_engine(engine, source, output, profile, spec, timeout):
    output.mkdir(parents=True,exist_ok=True); profile.mkdir(parents=True,exist_ok=True)
    # Private profile, no trusted macro locations; no user's Office profile is used.
    (profile/'registrymodifications.xcu').write_text(
        '<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
        '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
        '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop>'
        '<prop oor:name="SecureURL" oor:op="fuse"><value/></prop></item></oor:items>',encoding='utf-8')
    args=[str(engine),'-env:UserInstallation='+profile.as_uri(),'--headless','--norestore',
          '--nodefault','--nofirststartwizard','--convert-to',spec,'--outdir',str(output),str(source)]
    env=os.environ.copy(); env.update(TMPDIR='/private/tmp',TMP='/private/tmp',TEMP='/private/tmp')
    if os.uname().sysname == 'Darwin':
        from xml.sax.saxutils import escape
        fontdirs=[Path('/System/Library/Fonts'),Path('/System/Library/Fonts/Supplemental'),
                  Path('/Library/Fonts'),Path.home()/'Library/Fonts',
                  engine.parent.parent/'Resources/fonts',engine.parent.parent/'Resources/share/fonts']
        fc=profile/'fonts.conf'
        fc.write_text('<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd"><fontconfig>'
                      +''.join('<dir>'+escape(str(d))+'</dir>' for d in fontdirs if d.is_dir())
                      +'<cachedir>'+escape(str(profile/'font-cache'))+'</cachedir></fontconfig>',encoding='utf-8')
        env['FONTCONFIG_FILE']=str(fc)
    p=subprocess.Popen(args,stdout=subprocess.PIPE,stderr=subprocess.PIPE,env=env,start_new_session=True)
    try: stdout,stderr=p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        os.killpg(p.pid,signal.SIGKILL); p.communicate()
        raise ConversionError('转换超时；已停止本次任务，原件不变。')
    expected=output/(source.stem+'.'+spec.split(':',1)[0])
    if p.returncode!=0 or not expected.is_file() or expected.stat().st_size==0:
        raise ConversionError('转换引擎无法读取此文档，或文件已损坏 / 受保护。')
    return expected

def publish(data, directory, stem):
    directory.mkdir(parents=True,exist_ok=True)
    if not directory.is_dir(): raise ConversionError('输出位置不是文件夹。')
    for index in range(10000):
        name=f'{stem}（转换）'+(f' {index+1}' if index else '')+'.docx'
        target=directory/name
        try: fd=os.open(target,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o644)
        except FileExistsError: continue
        try:
            with os.fdopen(fd,'wb') as f: f.write(data); f.flush(); os.fsync(f.fileno())
        except BaseException:
            target.unlink(missing_ok=True); raise
        return target
    raise ConversionError('同名输出过多，请选择其他文件夹。')

def convert_one(source, directory, engine, timeout=90):
    source=Path(source).expanduser().resolve(); directory=Path(directory).expanduser().resolve()
    result={'input':str(source),'status':'error','output':None,'message':'','verification':{}}
    try:
        if not source.is_file(): raise ConversionError('输入不是普通文件。')
        if source.stat().st_size>MAX_SIZE: raise ConversionError('当前版本支持不超过 100 MB 的单个文件。')
        original=source.read_bytes(); original_hash=digest(original)
        kind,baseline=detect(original)
        verification={'detected_format':kind,'original_sha256':original_hash}
        warnings=[]
        if kind=='docx':
            converted=original
            if baseline['noncanonical_names']:
                import io
                z,names=zip_members(original); buf=io.BytesIO()
                try:
                    with ZipFile(buf,'w') as out:
                        for name,info in names.items():
                            clone=ZipInfo(name,info.date_time); clone.compress_type=info.compress_type
                            clone.external_attr=info.external_attr
                            out.writestr(clone,z.read(info))
                finally: z.close()
                converted=buf.getvalue()
                verification['content_check']='仅修复压缩包内部路径；所有内部文件内容保持原样'
            else: verification['content_check']='原文件字节完全保留'
            verification['byte_identical']=converted==original
            verification['paragraphs']=baseline['paragraphs']
        else:
            if not Path(engine).is_file(): raise ConversionError('未找到应用内的转换引擎。')
            with tempfile.TemporaryDirectory(prefix='DocConverter-',dir='/private/tmp') as temp:
                root=Path(temp); staged=root/('input.'+kind); staged.write_bytes(original)
                before=(staged if kind=='odt' else convert_engine(Path(engine),staged,root/'baseline',root/'profile1','odt:writer8',timeout))
                imported=odt_info(before)
                target=convert_engine(Path(engine),staged,root/'converted',root/'profile2','docx:Office Open XML Text',timeout)
                converted=target.read_bytes(); after=docx_info(converted)
                if imported['tables']!=after['tables'] or imported['images']!=after['images']:
                    raise ConversionError('转换前后表格或图片数量不一致，未输出可能缺失内容的文档。')
                if imported['image_hashes']!=after['image_hashes']:
                    raise ConversionError('转换前后图片数据不同，当前版本无法可靠核验，未输出。')
                # Independent Word piece-table text gate when available.
                source_text=baseline['main'] if kind=='doc' else imported['main']
                if source_text==after['main']:
                    verification['content_check']='正文逐字符一致（含空格和段落）'
                elif visible_binary(source_text)==visible_xml(after['main']):
                    verification['content_check']='正文文字一致；段落 / 表格控制标记按 DOCX 结构转换'
                else:
                    raise ConversionError('转换前后正文文字不同，未输出可能改变内容的文档。')
                verification.update(paragraphs=after['paragraphs'],tables=after['tables'],images=after['images'],
                                    source_reader=('独立 Word 二进制解析' if kind=='doc' else '转换引擎读取基准'),
                                    fonts=after['fonts'])
                if imported['objects']:
                    raise ConversionError('文档含嵌入对象，当前版本无法可靠核验对象内容，未输出。')
                if after['revisions']:
                    raise ConversionError('文档含修订记录，当前版本无法完整核验修订，未输出。')
                from content_audit import audit_extra
                extra=audit_extra(kind,baseline if kind=='doc' else None,before,target)
                if not extra['passed']: raise ConversionError(extra['message'])
                verification['extra_content']=extra['details']
                warnings.append('内容核对通过；请用 Word 检查实际分页和版式')
        if digest(source.read_bytes())!=original_hash:
            raise ConversionError('转换期间原文件被其他程序修改，本次不输出。')
        verification['original_unchanged']=True
        output=publish(converted,directory,source.stem)
        verification['output_sha256']=digest(converted)
        result.update(status='warning' if warnings else 'ok',output=str(output),verification=verification,
                      message='转换完成。'+('；'.join(warnings) if warnings else '内容原样保留。'))
    except Exception as e:
        result['message']=str(e) if isinstance(e,ConversionError) else '文件无法可靠读取，已保留原件；'+str(e)
    return result

def main():
    p=argparse.ArgumentParser(description='本地转换文字文档为 DOCX，始终保留原件。')
    p.add_argument('--engine',required=True); p.add_argument('--output-dir',required=True)
    p.add_argument('--timeout',type=int,default=90); p.add_argument('inputs',nargs='+')
    args=p.parse_args(); failed=False
    for source in args.inputs:
        print(json.dumps({'event':'start','input':str(Path(source).expanduser().resolve())},ensure_ascii=False),flush=True)
        result=convert_one(source,args.output_dir,args.engine,args.timeout)
        print(json.dumps(result,ensure_ascii=False),flush=True); failed |= result['status']=='error'
    return 1 if failed else 0

if __name__=='__main__': raise SystemExit(main())
