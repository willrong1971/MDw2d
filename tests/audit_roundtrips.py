from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from PIL import Image
import subprocess, json, hashlib, shutil, os, difflib, io, argparse
ROOT=Path(__file__).resolve().parent/'_work'/'fixtures'
AUDIT=ROOT/'audit-roundtrips'
SOFFICE=os.environ.get('LO_ENGINE', str(Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/MacOS/soffice'))
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; T='urn:oasis:names:tc:opendocument:xmlns:text:1.0'; O='urn:oasis:names:tc:opendocument:xmlns:office:1.0'; S='urn:oasis:names:tc:opendocument:xmlns:style:1.0'; D='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'; TB='urn:oasis:names:tc:opendocument:xmlns:table:1.0'; X='http://www.w3.org/1999/xlink'
N={'w':W,'text':T,'office':O,'style':S,'draw':D,'table':TB,'xlink':X}
def sha(b): return hashlib.sha256(b).hexdigest()
def odt_flat(e,exclude_notes=True):
    s=e.text or ''
    for c in e:
        if exclude_notes and c.tag in ['{%s}note'%T,'{%s}annotation'%O]: pass
        elif c.tag=='{%s}s'%T: s+=' '*int(c.get('{%s}c'%T,'1'))
        elif c.tag=='{%s}tab'%T: s+='\t'
        elif c.tag=='{%s}line-break'%T: s+='\n'
        elif c.tag=='{%s}page-number'%T: pass
        else: s+=odt_flat(c,exclude_notes)
        s+=c.tail or ''
    return s

def odt_pars(container,exclude_notes=True):
    pars=[]
    for p in container.xpath('.//text:p | .//text:h',namespaces=N):
        if exclude_notes and any(a.tag in ['{%s}note'%T,'{%s}annotation'%O] for a in p.iterancestors()): continue
        pars.append(odt_flat(p,exclude_notes))
    return pars

def img_info(z,names):
    result=[]
    for name in names:
        data=z.read(name)
        try:
            im=Image.open(io.BytesIO(data)).convert('RGBA')
            result.append({'dimensions':list(im.size),'pixel_sha256':sha(im.tobytes()),'file_sha256':sha(data)})
        except Exception as exc:
            result.append({'file_sha256':sha(data),'unreadable_as_bitmap':str(exc)})
    return result

def odt_info(p):
    with ZipFile(p) as z:
        content=etree.fromstring(z.read('content.xml')); styles=etree.fromstring(z.read('styles.xml'))
        body=content.find('office:body',N)
        images=[e.get('{%s}href'%X) for e in body.xpath('.//draw:image',namespaces=N)]
        data={'main':odt_pars(body), 'header':[], 'footer':[], 'footnote':[], 'comment':[], 'tables':len(body.xpath('.//table:table',namespaces=N)),'images':img_info(z,images)}
        for k in ['header','footer']:
            data[k]=[odt_pars(e) for e in styles.xpath('//style:'+k,namespaces=N)]
        data['footnote']=[odt_pars(e,False) for e in content.xpath('//text:note-body',namespaces=N)]
        data['comment']=[odt_pars(e,False) for e in content.xpath('//office:annotation',namespaces=N)]
        return data

def w_text(p):
    s=''; field=[]
    for e in p.iter():
        if e.tag=='{%s}fldSimple'%W:
            continue
        if e.tag=='{%s}fldChar'%W:
            k=e.get('{%s}fldCharType'%W)
            if k=='begin': field.append({'code':'','result':False})
            elif k=='separate' and field: field[-1]['result']=True
            elif k=='end' and field: field.pop()
            continue
        if e.tag=='{%s}instrText'%W and field: field[-1]['code']+=e.text or ''; continue
        if any(not f['result'] or f['code'].strip().upper().startswith('PAGE') for f in field): continue
        if any(a.tag=='{%s}fldSimple'%W and a.get('{%s}instr'%W,'').strip().upper().startswith('PAGE') for a in e.iterancestors()): continue
        if e.tag=='{%s}t'%W: s+=e.text or ''
        elif e.tag=='{%s}tab'%W and e.getparent() is not None and e.getparent().tag=='{%s}r'%W: s+='\t'
        elif e.tag=='{%s}br'%W: s+='\n'
    return s

def w_pars(root): return [w_text(p) for p in root.xpath('.//w:p',namespaces=N)]
def docx_info(p):
    with ZipFile(p) as z:
        names=z.namelist(); body=etree.fromstring(z.read('word/document.xml')).find('w:body',N)
        data={'main':w_pars(body),'header':[], 'footer':[], 'footnote':[], 'comment':[], 'tables':len(body.xpath('.//w:tbl',namespaces=N)), 'images':img_info(z,[n for n in names if n.startswith('word/media/')])}
        for k in ['header','footer']:
            data[k]=[w_pars(etree.fromstring(z.read(n))) for n in sorted(names) if n.startswith('word/'+k) and n.endswith('.xml')]
        if 'word/footnotes.xml' in names:
            f=etree.fromstring(z.read('word/footnotes.xml'))
            data['footnote']=[w_pars(e) for e in f.xpath('//w:footnote',namespaces=N) if e.get('{%s}type'%W) not in ['separator','continuationSeparator']]
        if 'word/comments.xml' in names:
            c=etree.fromstring(z.read('word/comments.xml')); data['comment']=[w_pars(e) for e in c.xpath('//w:comment',namespaces=N)]
        return data

def meaningful(d):
    # Compare textual content without paragraphs/layout-only line/page breaks;
    # preserve every ordinary space, NBSP, tab and ideographic space.
    out={}
    for key in ['main','header','footer','footnote','comment']:
        vals=d[key]
        if key=='main': out[key]=''.join(vals).replace('\n','')
        elif key in ['header','footer']: out[key]=sorted(set(''.join(v).replace('\n','') for v in vals if ''.join(v)))
        else: out[key]=sorted(''.join(v).replace('\n','') for v in vals)
    out['tables']=d['tables']
    out['images']=sorted((tuple(im.get('dimensions',[])),im.get('pixel_sha256')) for im in d['images'])
    return out

def main():
    global ROOT, AUDIT, SOFFICE
    parser=argparse.ArgumentParser(description='Audit synthetic format roundtrips.')
    parser.add_argument('--fixtures', type=Path, default=ROOT)
    parser.add_argument('--engine', default=SOFFICE)
    args=parser.parse_args()
    ROOT=args.fixtures.expanduser().resolve()
    AUDIT=ROOT/'audit-roundtrips'; AUDIT.mkdir(parents=True, exist_ok=True)
    SOFFICE=str(Path(args.engine).expanduser().resolve())
    output={}
    env=os.environ.copy(); env['FONTCONFIG_FILE']=str(ROOT/'fonts.conf')
    for fmt in ['doc','rtf','odt']:
        source=ROOT/('baseline.'+fmt); folder=AUDIT/fmt; folder.mkdir(exist_ok=True)
        staged=folder/('source.'+fmt); shutil.copyfile(source,staged)
        for target,filter_ in [('odt','odt:writer8'),('docx','docx:Office Open XML Text')]:
            if fmt==target:
                continue
            dest=folder/target; dest.mkdir(exist_ok=True)
            cmd=[SOFFICE,'-env:UserInstallation='+(folder/('profile-'+target)).as_uri(),'--headless','--convert-to',filter_,'--outdir',str(dest),str(staged)]
            r=subprocess.run(cmd,capture_output=True,text=True,timeout=60,env=env)
            if r.returncode or not (dest/('source.'+target)).exists(): raise RuntimeError(r.stderr+r.stdout)
        original_odt=staged if fmt=='odt' else folder/'odt/source.odt'
        converted_docx=folder/'docx/source.docx'
        before=odt_info(original_odt); after=docx_info(converted_docx)
        b=meaningful(before); a=meaningful(after)
        diff={k:{'before':b[k],'after':a[k]} for k in b if b[k]!=a[k]}
        output[fmt]={'source_sha256':sha(source.read_bytes()),'source_odt_view':str(original_odt),'output_docx_view':str(converted_docx),'source_features':before,'output_features':after,'semantic_differences':diff,'matches':not diff,'layout_part_count_differences':{k:{'source':len(before[k]),'output':len(after[k])} for k in ['header','footer'] if len(before[k])!=len(after[k])}}
    (AUDIT/'roundtrip-results.json').write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({fmt:{'matches':r['matches'],'difference_parts':list(r['semantic_differences'])} for fmt,r in output.items()},ensure_ascii=False,indent=2))
if __name__=='__main__': main()
