pyinstaller --distpath ./ --noconfirm --log-level WARN ^
    -F -n GOC --nowindowed ^
    --add-data="Lib\site-packages\tinycss2\VERSION;." ^
    --hidden-import ezdxf ^
    --hidden-import MainProgramDefinitions ^
    --hidden-import pkg_resources.py2_warn ^
    --hidden-import pyi_rth_pkgres ^
    --hidden-import svgwrite -d bootloader^
    --i logo.ico --runtime-tmpdir ./ --win-private-assemblies^
    convertor_1_1_3.py