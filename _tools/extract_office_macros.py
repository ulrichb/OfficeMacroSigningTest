import sys
from pathlib import Path
import re
import traceback
from oletools.olevba import VBA_Parser

def main():
    print('== Extract Office Macros ==')

    files_to_extract = Path('.').glob('**/*.xlsm')

    extraction_errors = 0

    for file_to_extract in files_to_extract:

        try:
            print(f'Extracting {file_to_extract} ...')

            extract_file(file_to_extract)

        except:
            print(f'Error while processing "{file_to_extract}":')
            print(traceback.format_exc())
            extraction_errors += 1

    return extraction_errors

def extract_file(file: Path):
    vbaparser = VBA_Parser(file)

    if not vbaparser.detect_vba_macros():
        raise Exception("File doesn't contain macros")
    
    assert not vbaparser.is_encrypted

    for (filename, _, vba_filename, vba_code) in vbaparser.extract_macros():

        extract_dir = Path(file.parent, f'{file.stem}_Macros')
        extract_dir.mkdir(parents=True, exist_ok=True)

        extract_file_path = Path(extract_dir, f'{vba_filename}.vb')

        assert filename == 'xl/vbaProject.bin', f'Unexpected {filename}'

        print(f'Extracting "{filename}/{vba_filename}" into "{extract_file_path}" ...')

        vba_code = remove_standard_attribute_preamble_in_vba_code(filename=vba_filename, vba_code=vba_code)

        extract_file_path.write_text(vba_code)

    vbaparser.close()

def remove_standard_attribute_preamble_in_vba_code(filename: str, vba_code: str):
    vba_code = re.sub(r'^Attribute VB_Name = "' + re.escape(filename) + '"\r?\n', '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r'^Attribute VB_Base = "0(\{[A-F0-9-]{36}\})+"\r?\n', '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_GlobalNameSpace = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_Creatable = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_PredeclaredId = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_Exposed = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_TemplateDerived = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    vba_code = re.sub(r"^Attribute VB_Customizable = (False|True)\r?\n", '', vba_code, flags=re.MULTILINE)
    return vba_code

if __name__ == '__main__':
    sys.exit(main())
