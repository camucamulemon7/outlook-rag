"""Feature modules must load without starting the application or Windows COM."""
import json
import os
from pathlib import Path
import subprocess
import sys
import textwrap
import unittest


class IndependentModuleImports(unittest.TestCase):
    def test_feature_modules_do_not_load_facade_or_windows_apis(self):
        root = Path(__file__).resolve().parents[1]
        code = textwrap.dedent('''
            import importlib
            import importlib.abc
            import json
            import sys

            forbidden = {'outlook_rag.app', 'pythoncom', 'win32com.client', 'win32crypt'}
            class Guard(importlib.abc.MetaPathFinder):
                def find_spec(self, fullname, path, target=None):
                    if fullname in forbidden:
                        raise AssertionError('Unexpected runtime initialization: ' + fullname)
            sys.meta_path.insert(0, Guard())
            modules = ['settings', 'embeddings', 'indexing', 'sync', 'maintenance', 'database', 'features']
            for name in modules:
                importlib.import_module('outlook_rag.' + name)
            assert not forbidden.intersection(sys.modules)
            print(json.dumps({'imported': modules}))
        ''')
        env = {k: v for k, v in os.environ.items() if not k.startswith('OUTLOOK_RAG_')}
        env['PYTHONPATH'] = str(root / 'src')
        result = subprocess.run([sys.executable, '-c', code], env=env,
                                capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(json.loads(result.stdout)['imported']), 7)


if __name__ == '__main__':
    unittest.main()
