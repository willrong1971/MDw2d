from pathlib import Path
import hashlib, importlib.util, json, sys, argparse
SCRIPT_DIR=Path(__file__).resolve().parent
WORK=SCRIPT_DIR/'_work'
ROOT=WORK/'fixtures'
parser=argparse.ArgumentParser()
parser.add_argument('--fixtures',type=Path,default=ROOT)
parser.add_argument('--results',type=Path,default=WORK/'general-test-results.ndjson')
parser.add_argument('--report',type=Path,default=WORK/'general-test-audit.json')
args=parser.parse_args()
ROOT=args.fixtures.expanduser().resolve()
s=importlib.util.spec_from_file_location('roundtrip_audit',SCRIPT_DIR/'audit_roundtrips.py'); helper=importlib.util.module_from_spec(s); s.loader.exec_module(helper)
manifest=json.loads((ROOT/'expectations.json').read_text())
roundtrips=json.loads((ROOT/'audit-roundtrips/roundtrip-results.json').read_text())
results={}
for line in args.results.read_text().splitlines():
 try: row=json.loads(line)
 except json.JSONDecodeError: continue
 if 'status' in row: results[Path(row['input']).name]=row
checks=[]
for exp in manifest['fixtures']:
 name=exp['filename']; source=ROOT/name; row=results.get(name)
 sha=hashlib.sha256(source.read_bytes()).hexdigest()
 item={'filename':name,'expected_success':exp['expected_success'],'original_unchanged':sha==manifest['sha256'][name],'source_sha256':sha,'result':row}
 if row is None:
  item['status_matches_expectation']=False; item['missing_result']=True
 else:
  success=row['status'] in ['ok','warning'] and row.get('output') is not None
  item['status_matches_expectation']=success==exp['expected_success']
  if success:
   output=Path(row['output']); raw=output.read_bytes(); item['output_sha256']=hashlib.sha256(raw).hexdigest()
   if exp.get('byte_identical_expected'):
    item['byte_identical']=raw==source.read_bytes()
    before=helper.docx_info(source)
   else:
    family={'ole-word':'doc','rtf':'rtf','odt':'odt'}[exp['detected_family']]
    before=helper.odt_info(source if family=='odt' else Path(roundtrips[family]['source_odt_view']))
   after=helper.docx_info(output); b=helper.meaningful(before); a=helper.meaningful(after)
   item['semantic_differences']={k:{'source':b[k],'output':a[k]} for k in b if b[k]!=a[k]}
   item['fidelity_checks_pass']=not item['semantic_differences'] and (not exp.get('byte_identical_expected') or item['byte_identical'])
   item['source_features']=before; item['output_features']=after
   item['layout_part_count_differences']={k:{'source':len(before[k]),'output':len(after[k])} for k in ['header','footer'] if len(before[k])!=len(after[k])}
 checks.append(item)
report={'independent_audit':True,'comparison_rules':['Source format is compared with its own imported ODT baseline; losses introduced when the synthetic fixture was exported to DOC/RTF/ODT are not charged to this conversion.','Body excludes note/comment contents; they are compared independently.','Every ordinary space, NBSP, ideographic space and actual run tab is retained; paragraph and page/line break markers are excluded from flattened content checks.','Paragraph tab-stop definitions are formatting and are not treated as actual TAB characters.','Header/footer unique nonempty text is compared because duplicate first/even/default parts may represent the same source content.','Images are compared by decoded RGBA pixel hash and dimensions.'],'summary':{'fixture_count':len(checks),'results_present':len(results),'original_hashes_unchanged':all(i['original_unchanged'] for i in checks),'expectation_matches':sum(i['status_matches_expectation'] for i in checks),'successful_results':sum(i.get('fidelity_checks_pass',False) for i in checks),'unmet_expected_success':[i['filename'] for i in checks if i['expected_success'] and not i['status_matches_expectation']],'fidelity_failures':[i['filename'] for i in checks if i.get('fidelity_checks_pass') is False]},'fixtures':checks,'direct_roundtrip_control_experiments':roundtrips}
args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
print(json.dumps(report['summary'],ensure_ascii=False,indent=2))
for i in checks:
 print(i['filename'],i['result']['status'] if i.get('result') else 'missing','hash='+str(i['original_unchanged']),'fidelity='+str(i.get('fidelity_checks_pass','n/a')))

if not (report['summary']['original_hashes_unchanged'] and report['summary']['expectation_matches']==report['summary']['fixture_count'] and not report['summary']['fidelity_failures']):
 raise SystemExit(1)
