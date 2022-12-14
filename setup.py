from setuptools import Extension, setup, find_packages
from setuptools.command.build_ext import build_ext
from distutils.unixccompiler import UnixCCompiler

class custom_build_ext(build_ext):
    def initialize_options(self) -> None:
        UnixCCompiler.src_extensions.append(".mm")
        UnixCCompiler.language_map[".mm"] = "objc"

        og_compile = UnixCCompiler._compile
        og_link = UnixCCompiler.link

        def new_compile(self, obj, src, ext, cc_args, extra_postargs, pp_opts):
            new_postargs = extra_postargs
            if ext == ".cpp":
                new_postargs = new_postargs + ["-std=c++17"]
            elif ext == ".mm":
                new_postargs = new_postargs + [
                    "-ObjC++", "-fobjc-weak", "-fobjc-arc"
                ]
            og_compile(self, obj, src, ext, cc_args, extra_postargs, pp_opts)

        def new_link(self,
                     target_desc,
                     objects,
                     output_filename,
                     output_dir=None,
                     libraries=None,
                     library_dirs=None,
                     runtime_library_dirs=None,
                     export_symbols=None,
                     debug=False,
                     extra_preargs=None,
                     extra_postargs=None,
                     build_temp=None,
                     target_lang=None):
            new_extra_postargs = extra_postargs or []
            framework_postargs = [
                "-framework", "AppKit",
                "-shared"
            ]

            new_extra_postargs = new_extra_postargs + framework_postargs

            og_link(self, target_desc, objects, output_filename, output_dir,
                    libraries, library_dirs, runtime_library_dirs, export_symbols, debug,
                    extra_preargs, new_extra_postargs, build_temp, target_lang)

        UnixCCompiler._compile = new_compile
        UnixCCompiler.link = new_link
        super().initialize_options()

    def build_extensions(self) -> None:
        self.compiler.set_executable("compiler_so", "clang++")
        self.compiler.set_executable("compiler_cxx", "clang++")
        self.compiler.set_executable("linker_so", "clang++")
        build_ext.build_extensions(self)


setup(
    name="officecharts",
    version="1.0",
    description="Create Office Graphics Charts from Python",
    author="Michael Allen",
    author_email="michael@velofrog.com",
    packages=find_packages(),
    setup_requires=['pandas', 'python-dateutil', 'XlsxWriter', 'matplotlib'],
    ext_modules=[
        Extension(
            name="officecharts.os_clipboard",
            sources=["src/mac_clipboard.mm"]
        )
    ],
    cmdclass={"build_ext": custom_build_ext}
)