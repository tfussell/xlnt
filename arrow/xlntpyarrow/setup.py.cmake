from distutils.core import setup, Extension
from distutils import sysconfig
import os

description = """
xlntpyarrow allows Apache Arrow tables to be written to and read from an XLSX
file efficiently using the C++ library xlnt.
""".strip()

cfg_vars = sysconfig.get_config_vars()
if 'CFLAGS' in cfg_vars:
    cfg_vars['CFLAGS'] = cfg_vars['CFLAGS'].replace('-Wstrict-prototypes', '')

project_root = '${CMAKE_SOURCE_DIR}'
conda_root = '${ARROW_INCLUDE_PATH}'

include_dirs = [
    os.path.join(project_root, 'include'),
    os.path.join(project_root, 'arrow/xlntarrow'),
    os.path.join(project_root, 'arrow/xlntpyarrow'),
    conda_root
]

library_dirs = [
    os.path.join(project_root, 'build/arrow/xlntarrow/Debug'),
    os.path.join(project_root, 'build/source/Debug'),
    os.path.join(conda_root, '../lib')
]

xlntpyarrow_extension = Extension(
    'xlntpyarrow',
    ['${CMAKE_CURRENT_SOURCE_DIR}/xlntpyarrow.cpp'],
    language = 'c++',
    include_dirs = include_dirs,
    libraries = [
        'arrow',
        'xlntarrow',
        'xlnt'
    ],
    library_dirs = library_dirs,
    extra_compile_args=['${CMAKE_CXX_FLAGS}']
)

classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Environment :: Plugins',
    'Intended Audience :: Science/Research',
    'License :: OSI Approved :: MIT License',
    'Natural Language :: English',
    'Operating System :: Microsoft :: Windows',
    'Operating System :: MacOS :: MacOS X',
    'Operating System :: POSIX :: Linux',
    'Programming Language :: C',
    'Programming Language :: C++',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3.6',
    'Programming Language :: Python :: Implementation :: CPython',
    'Topic :: Database',
    'Topic :: Office/Business :: Financial :: Spreadsheet',
    'Topic :: Scientific/Engineering :: Information Analysis',
    'Topic :: Software Development :: Libraries :: Python Modules'
]

setup(
    name = 'xlntpyarrow',
    version = '1.1.0',
    classifiers = classifiers,
    description = description,
    ext_modules = [xlntpyarrow_extension],
    author = 'Thomas Fussell',
    author_email = 'thomas.fussell@gmail.com',
    url = 'https://github.com/tfussell/xlnt'
)
