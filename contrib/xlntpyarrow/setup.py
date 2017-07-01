from distutils.core import setup, Extension

setup(name='xlntpyarrow',
    include_dirs = [
        'C:/Users/Thomas/Development/xlnt/contrib/xlnt_arrow',
        'C:/Users/Thomas/Development/xlnt/contrib/xlntpyarrow',
        'C:/Users/Thomas/Miniconda3/Library/include'],
    ext_modules=[Extension("xlntpyarrow", ["xlntpyarrow.cpp"],
        libraries = [
            'arrow',
            'xlnt_arrow',
            'xlnt'
        ],
        library_dirs = [
            'C:/Users/Thomas/Miniconda3/Library/lib',
            'C:/Users/Thomas/Development/xlnt/build/contrib/xlnt_arrow/Release',
            'C:/Users/Thomas/Development/xlnt/build/source/Release'
        ])]
)
