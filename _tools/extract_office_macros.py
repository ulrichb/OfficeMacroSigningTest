import sys
from pathlib import Path
import traceback
import glob
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

        extract_file_path.write_text(vba_code)

    vbaparser.close()

if __name__ == '__main__':
    sys.exit(main())
