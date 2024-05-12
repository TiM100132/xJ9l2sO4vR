import os
from zipfile import ZipFile
import shutil
import tempfile

class FileHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.temp_dir = None

    def extract_contents(self):
        xml_file_paths = []
        try:
            self.temp_dir = tempfile.mkdtemp()

            with ZipFile(self.file_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)

            sheets_dir = os.path.join(self.temp_dir, 'xl', 'worksheets')
            
            for filename in os.listdir(sheets_dir):
                if filename.endswith('.xml'):
                    xml_file_paths.append(os.path.join(sheets_dir, filename))
     
            return xml_file_paths, self.temp_dir
    
        except Exception as e:
            print(f"Ошибка при обработке XLSX файла: {e}")

    def del_tmp_dir(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)
