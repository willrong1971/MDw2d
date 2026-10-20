#!/usr/bin/env python3
"""Batch CLI parsing and dispatch; all document paths and content are synthetic."""
from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
import converter


class BatchInputsTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='mdw2d-batch-test-')
        self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name).resolve()
        self.manifest=self.root/'待转换 清单.json'
        self.output=self.root/'集中 输出'
        self.engine=self.root/'unused-engine'
        self.base=['--engine',str(self.engine)]

    def write_manifest(self,value):
        self.manifest.write_text(json.dumps(value,ensure_ascii=False),encoding='utf-8')
        return ['--input-list-file',str(self.manifest)]

    @staticmethod
    def converted(source,directory,engine,timeout):
        return {'input':str(Path(source).expanduser().resolve()),'status':'ok',
                'output':str(Path(directory)/'result.docx'),'message':'ok','verification':{}}

    def run_cli(self,argv,side_effect=None):
        output=io.StringIO()
        with patch.object(converter,'convert_one',side_effect=side_effect or self.converted) as convert:
            with redirect_stdout(output):
                status=converter.main(self.base+argv)
        events=[json.loads(line) for line in output.getvalue().splitlines()]
        return status,events,convert

    def assert_parse_error(self,argv,expected):
        stdout,stderr=io.StringIO(),io.StringIO()
        with patch.object(converter,'convert_one') as convert:
            with redirect_stdout(stdout),redirect_stderr(stderr):
                with self.assertRaises(SystemExit) as raised:
                    converter.main(self.base+argv)
        self.assertEqual(raised.exception.code,2)
        self.assertIn(expected,stderr.getvalue())
        self.assertEqual(stdout.getvalue(),'')
        convert.assert_not_called()
        self.assertFalse(self.output.exists())

    def test_legacy_positional_and_central_output(self):
        sources=[str(self.root/'第一份 空格.wps'),str(self.root/'另一个.doc')]
        code,events,convert=self.run_cli(['--output-dir',str(self.output),*sources])
        self.assertEqual(code,0)
        self.assertEqual([call.args for call in convert.call_args_list],
                         [(source,str(self.output),str(self.engine),90) for source in sources])
        self.assertEqual(len(events),4)
        for index,source in enumerate(sources):
            self.assertEqual(events[index*2],{'event':'start','input':str(Path(source).resolve())})
            self.assertEqual(events[index*2+1]['status'],'ok')

    def test_manifest_beside_each_source(self):
        sources=[str(self.root/'子目录 甲'/'文档一.wps'),str(self.root/'乙'/'文档二.doc')]
        code,events,convert=self.run_cli(['--beside-input',*self.write_manifest(sources)])
        self.assertEqual(code,0)
        self.assertEqual([call.args[1] for call in convert.call_args_list],
                         [Path(source).parent for source in sources])
        self.assertEqual([event['input'] for event in events[::2]],sources)

    def test_manifest_central_output_and_timeout(self):
        source=str(self.root/'文档.rtf')
        code,_,convert=self.run_cli(['--output-dir',str(self.output),'--timeout','23',
                                     *self.write_manifest([source])])
        self.assertEqual(code,0)
        convert.assert_called_once_with(source,str(self.output),str(self.engine),23)

    def test_relative_manifest_paths_are_relative_to_cwd(self):
        sources=['子目录 甲/文件.wps','第二.doc']
        with patch('os.getcwd',return_value=str(self.root)):
            code,events,convert=self.run_cli(['--beside-input',*self.write_manifest(sources)])
        self.assertEqual(code,0)
        self.assertEqual([event['input'] for event in events[::2]],
                         [str(self.root/source) for source in sources])
        self.assertEqual([call.args[1] for call in convert.call_args_list],
                         [(self.root/source).parent for source in sources])

    def test_failure_does_not_stop_later_files(self):
        sources=[str(self.root/'损坏.doc'),str(self.root/'正常.rtf')]
        def result(source,directory,engine,timeout):
            value=self.converted(source,directory,engine,timeout)
            if source==sources[0]:
                value.update(status='error',output=None,message='synthetic invalid file')
            return value
        code,events,convert=self.run_cli(['--beside-input',*self.write_manifest(sources)],result)
        self.assertEqual(code,1)
        self.assertEqual(convert.call_count,2)
        self.assertEqual([event['status'] for event in events[1::2]],['error','ok'])

    def test_symlink_loop_does_not_stop_later_files(self):
        loop=self.root/'loop.doc'
        loop.symlink_to(loop.name)
        later=self.root/'正常.rtf'
        code,events,convert=self.run_cli(['--beside-input',*self.write_manifest([str(loop),str(later)])])
        self.assertEqual(code,1)
        self.assertEqual([event['status'] for event in events[1::2]],['error','ok'])
        convert.assert_called_once_with(str(later),self.root,str(self.engine),90)

    def test_unexpected_per_file_exception_does_not_stop_later_files(self):
        sources=[str(self.root/'第一份.doc'),str(self.root/'第二份.rtf')]
        later=self.converted(sources[1],self.output,self.engine,90)
        code,events,convert=self.run_cli(['--output-dir',str(self.output),*sources],
                                        [RuntimeError('synthetic path-resolution failure'),later])
        self.assertEqual(code,1)
        self.assertEqual(convert.call_count,2)
        self.assertEqual([event['status'] for event in events[1::2]],['error','ok'])
        self.assertIn('synthetic path-resolution failure',events[1]['message'])

    def test_real_output_directory_symlink_loop_is_reported_for_every_file(self):
        output=self.root/'output-loop'; output.symlink_to(output.name)
        sources=[self.root/'first.rtf',self.root/'second.rtf']
        for source in sources:
            source.write_bytes(b'{\\rtf1\\ansi Synthetic input remains unchanged.}')
        originals={source:source.read_bytes() for source in sources}
        stdout=io.StringIO()
        with redirect_stdout(stdout):
            code=converter.main(self.base+['--output-dir',str(output),*[str(source) for source in sources]])
        events=[json.loads(line) for line in stdout.getvalue().splitlines()]
        self.assertEqual(code,1)
        self.assertEqual(len(events),4)
        self.assertEqual([event['status'] for event in events[1::2]],['error','error'])
        self.assertTrue(all(event['output'] is None for event in events[1::2]))
        self.assertEqual({source:source.read_bytes() for source in sources},originals)

    def test_large_manifest_exceeds_argv_limit(self):
        arg_max=os.sysconf('SC_ARG_MAX')
        source_prefix=str(self.root/('长路径 空格'*24))
        count=arg_max//len(source_prefix.encode('utf-8'))+30
        sources=[f'{source_prefix}{index:06}.doc' for index in range(count)]
        argv=self.write_manifest(sources)
        self.assertGreater(self.manifest.stat().st_size,arg_max)
        self.assertLess(sum(len(item.encode('utf-8')) for item in self.base+argv),4096)
        code,events,convert=self.run_cli(['--beside-input',*argv])
        self.assertEqual(code,0)
        self.assertEqual(convert.call_count,count)
        self.assertEqual(len(events),count*2)

    def test_invalid_schema_is_rejected_before_any_conversion(self):
        invalid=[None,{},'a.doc',3,True,[],['valid.doc',None],['valid.doc',8],
                 ['valid.doc',{}],['valid.doc',[]],['valid.doc',''],
                 ['valid.doc','  '],['valid.doc','bad\0.doc'],['valid.doc','bad\ud800.doc']]
        for value in invalid:
            with self.subTest(value=repr(value)):
                # ensure_ascii also allows an invalid surrogate JSON fixture.
                self.manifest.write_text(json.dumps(value,ensure_ascii=True),encoding='utf-8')
                self.assert_parse_error(['--output-dir',str(self.output),'--input-list-file',str(self.manifest)],
                                        'error:')

    def test_invalid_json_and_encoding_are_rejected(self):
        for value in [b'["ok.doc",',b'["ok.doc",NaN]',b'\xff\xfe',b'[[[[[']:
            with self.subTest(value=value):
                self.manifest.write_bytes(value)
                self.assert_parse_error(['--beside-input','--input-list-file',str(self.manifest)],'error:')

    def test_missing_manifest_is_rejected(self):
        self.assert_parse_error(['--beside-input','--input-list-file',str(self.manifest)],'无法读取输入清单')

    def test_modes_must_be_explicit_and_exclusive(self):
        for argv in [['input.doc'],['--beside-input','--output-dir',str(self.output),'input.doc'],
                     ['--beside-input'],['--output-dir',str(self.output)]]:
            with self.subTest(argv=argv):
                self.assert_parse_error(argv,'error:')

    def test_manifest_and_positional_inputs_cannot_mix(self):
        self.assert_parse_error(['--beside-input',*self.write_manifest(['list.doc']),'positional.doc'],
                                '不能与直接传入的文件路径同时使用')


if __name__=='__main__':
    unittest.main()
