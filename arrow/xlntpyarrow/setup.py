from distutils.core import setup, Extension, sysconfig

description = """
xlntpyarrow allows Apache Arrow tables to be written to and read from an XLSX
file efficiently using the C++ library xlnt.
""".strip()

cfg_vars = sysconfig.get_config_vars()
if 'CFLAGS' in cfg_vars:
    cfg_vars['CFLAGS'] = cfg_vars['CFLAGS'].replace('-Wstrict-prototypes', '')

xlntpyarrow_extension = Extension(
    'xlntpyarrow',
    ['xlntpyarrow.cpp'],
    language = 'c++',
    include_dirs = [
        '/root/xlnt/arrow/xlntarrow',
        '/root/xlnt/arrow/xlntpyarrow',
        '/root/miniconda3/include'
    ],
    libraries = [
        'arrow',
        'xlntarrow',
        'xlnt'
    ],
    library_dirs = [
        '/root/miniconda3/lib',
        '/root/xlnt/build/arrow/xlntarrow',
        '/root/xlnt/build/source'
    ],
    extra_compile_args=['-std=c++11']
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
