from distutils.core import setup, Extension

setup(name='xlntpyarrow',
    include_dirs = [
        'C:/Users/Thomas/Development/xlnt/arrow/xlntarrow',
        'C:/Users/Thomas/Development/xlnt/arrow/xlntpyarrow',
        'C:/Users/Thomas/Miniconda3/Library/include'],
    ext_modules=[Extension("xlntpyarrow", ["xlntpyarrow.cpp"],
        libraries = [
            'arrow',
            'xlntarrow',
            'xlnt'
        ],
        library_dirs = [
            'C:/Users/Thomas/Miniconda3/Library/lib',
            'C:/Users/Thomas/Development/xlnt/build/arrow/xlntarrow/Release',
            'C:/Users/Thomas/Development/xlnt/build/source/Release'
        ])]
)
