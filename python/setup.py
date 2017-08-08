import glob
import os
import os.path as osp
import re
import shutil
import sys

from distutils.command.clean import clean as _clean
from distutils import sysconfig

from os.path import join as pjoin

from setuptools import setup, Extension, Distribution
from setuptools.command.build_ext import build_ext as _build_ext

# Check if we're running 64-bit Python
is_64_bit = sys.maxsize > 2**32

class clean(_clean):
    def run(self):
        _clean.run(self)
        for x in []:
            try:
                os.remove(x)
            except OSError:
                pass

class build_ext(_build_ext):
    def run(self):
        self._run_cmake()

    # adapted from cmake_build_ext in dynd-python
    # github.com/libdynd/dynd-python

    description = "Build the C-extensions for arrow"
    user_options = ([('extra-cmake-args=', None, 'extra arguments for CMake'),
                     ('build-type=', None, 'build type (Debug or Release)')] +
                    _build_ext.user_options)

    def initialize_options(self):
        _build_ext.initialize_options(self)
        self.extra_cmake_args = os.environ.get('XLNTPYARROW_CMAKE_OPTIONS', '')
        self.build_type = os.environ.get('XLNTPYARROW_BUILD_TYPE', 'Debug')

        self.cmake_cxxflags = os.environ.get('XLNTPYARROW_CXXFLAGS', '')

        if sys.platform == 'win32':
            # Cannot do debug builds in Windows unless Python itself is a debug
            # build
            if not hasattr(sys, 'gettotalrefcount'):
                self.build_type = 'Release'

    def _run_cmake(self):
        # The directory containing this setup.py
        source = osp.dirname(osp.abspath(__file__))

        # The staging directory for the module being built
        build_temp = pjoin(os.getcwd(), self.build_temp)
        build_lib = os.path.join(os.getcwd(), self.build_lib)

        # Change to the build directory
        saved_cwd = os.getcwd()
        if not os.path.isdir(self.build_temp):
            self.mkpath(self.build_temp)
        os.chdir(self.build_temp)

        # Detect if we built elsewhere
        if os.path.isfile('CMakeCache.txt'):
            cachefile = open('CMakeCache.txt', 'r')
            cachedir = re.search('CMAKE_CACHEFILE_DIR:INTERNAL=(.*)',
                                 cachefile.read()).group(1)
            cachefile.close()
            if (cachedir != build_temp):
                return

        static_lib_option = ''

        cmake_options = [
            '-DPYTHON_EXECUTABLE=%s' % sys.executable,
            static_lib_option,
        ]

        if len(self.cmake_cxxflags) > 0:
            cmake_options.append('-DPYARROW_CXXFLAGS="{0}"'
                                 .format(self.cmake_cxxflags))

        cmake_options.append('-DCMAKE_BUILD_TYPE={0}'
                             .format(self.build_type))

        if 'CMAKE_GENERATOR' in os.environ:
            cmake_options.append('-G{}'
                .format(os.environ['CMAKE_GENERATOR']))

        if sys.platform != 'win32':
            cmake_options.append('-DCMAKE_INSTALL_PREFIX={0}'
                                 .format(os.environ['PREFIX']))
            cmake_command = (['cmake', self.extra_cmake_args] +
                             cmake_options + [source])

            print("-- Runnning cmake for xlntpyarrow")
            self.spawn(cmake_command)
            print("-- Finished cmake for xlntpyarrow")
            args = ['make']
            if os.environ.get('XLNTPYARROW_BUILD_VERBOSE', '0') == '1':
                args.append('VERBOSE=1')

            if 'XLNTPYARROW_PARALLEL' in os.environ:
                args.append('-j{0}'.format(os.environ['XLNTPYARROW_PARALLEL']))

            args.append('install')

            print("-- Running cmake --build for xlntpyarrow")
            self.spawn(args)
            print("-- Finished cmake --build for xlntpyarrow")
        else:
            import shlex
            if not is_64_bit:
                raise RuntimeError('Not supported on 32-bit Windows')

            cmake_options.append('-DCMAKE_INSTALL_PREFIX={0}'
                                 .format(os.environ['LIBRARY_PREFIX']))
            extra_cmake_args = shlex.split(self.extra_cmake_args)
            cmake_command = (['cmake'] + extra_cmake_args +
                             cmake_options + [source])

            print("-- Runnning cmake for xlntpyarrow")
            self.spawn(cmake_command)
            print("-- Finished cmake for xlntpyarrow")
            # Do the build
            print("-- Running cmake --build for xlntpyarrow")
            self.spawn(['cmake', '--build', '.', '--config', self.build_type, '--target', 'INSTALL'])
            print("-- Finished cmake --build for xlntpyarrow")

        if self.inplace:
            # a bit hacky
            build_lib = saved_cwd

        # Move the libraries to the place expected by the Python
        # build
        shared_library_prefix = 'lib'
        if sys.platform == 'darwin':
            shared_library_suffix = '.dylib'
        elif sys.platform == 'win32':
            shared_library_suffix = '.dll'
            shared_library_prefix = ''
        else:
            shared_library_suffix = '.so'

        try:
            os.makedirs(pjoin(build_lib, 'xlntpyarrow'))
        except OSError:
            pass

        self._found_names = []
        built_path = self.get_ext_built('lib')
        if not os.path.exists(built_path):
            raise RuntimeError('xlntpyarrow C-extension failed to build:',
                               os.path.abspath(built_path))

        ext_path = pjoin(build_lib, self._get_cmake_ext_path('lib'))
        if os.path.exists(ext_path):
            os.remove(ext_path)
        self.mkpath(os.path.dirname(ext_path))
        print('Moving built C-extension', built_path,
              'to build path', ext_path)
        shutil.move(self.get_ext_built('lib'), ext_path)
        self._found_names.append('lib')

        os.chdir(saved_cwd)

    def _failure_permitted(self, name):
        return False

    def _get_inplace_dir(self):
        pass

    def _get_cmake_ext_path(self, name):
        # Get the package directory from build_py
        build_py = self.get_finalized_command('build_py')
        package_dir = build_py.get_package_dir('xlntpyarrow')
        # This is the name of the arrow C-extension
        suffix = sysconfig.get_config_var('EXT_SUFFIX')
        if suffix is None:
            suffix = sysconfig.get_config_var('SO')
        filename = name + suffix
        return pjoin(package_dir, filename)

    def get_ext_built(self, name):
        if sys.platform == 'win32':
            head, tail = os.path.split(name)
            suffix = sysconfig.get_config_var('SO')
            return pjoin(head, self.build_type, tail + suffix)
        else:
            suffix = sysconfig.get_config_var('SO')
            print('suffix is', suffix)
            return name + suffix

    def get_names(self):
        return self._found_names

    def get_outputs(self):
        # Just the C extensions
        # regular_exts = _build_ext.get_outputs(self)
        return [self._get_cmake_ext_path(name)
                for name in self.get_names()]

long_description = """
xlntpyarrow allows Apache Arrow tables to be written to and read from an XLSX
file efficiently using the C++ library xlnt.
""".strip()

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

class BinaryDistribution(Distribution):
    def has_ext_modules(foo):
        return True

setup(
    name = 'xlntpyarrow',
    version = '1.1.0',
    description = 'Python library for converting Apache Arrow tables<->XLSX files',
    long_description = long_description,
    author = 'Thomas Fussell',
    author_email = 'thomas.fussell@gmail.com',
    maintainer = 'Thomas Fussell',
    maintainer_email = 'thomas.fussell@gmail.com',
    url = 'https://github.com/tfussell/xlnt',
    download_url = 'https://github.com/tfussell/xlnt/releases',
    packages = ['xlntpyarrow'],
    ext_modules = [Extension('xlntpyarrow.lib', [])],
    cmdclass = {
        'clean': clean,
        'build_ext': build_ext
    },
    classifiers = classifiers,
    distclass = BinaryDistribution,
    zip_safe = False
)
