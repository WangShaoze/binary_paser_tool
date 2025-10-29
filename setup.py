from distutils.core import setup
import py2exe

setup(
    windows=[{
        'script': 'test-02.py',
        'dest_base': '二进制数据解析工具',
        'uac_info': "requireAdministrator"
    }],
    options={
        'py2exe': {
            'bundle_files': 1,
            'compressed': True,
            'includes': ['openpyxl', 'ctypes', 'base64', 'os', 'tkinter', 'datetime'],
            'excludes': ['_gtkagg', '_tkagg', 'cairo', 'pango', 'numeric', 'pygtk', 'scipy']
        }
    },
    zipfile=None
)