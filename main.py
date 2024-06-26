import os
import sys
from file_handler import FileHandler
from xml_splitter import XMLSplitter
from sheet_cells_parser import SheetCellsParser
from extract_ooxml_object_info import ExtractObjectInfo

class Analyzer:
    def __init__(self, file_path):
        self.file_handler = FileHandler(file_path)
        self.xml_content, self.temp_dir = self.file_handler.extract_contents()  # Список XML файлов листов
        self.path_for_test_file = os.path.join(os.path.dirname(file_path), "test_file.txt")
        self.extract_objects = ExtractObjectInfo(self.temp_dir)
        
        if os.path.exists(getattr(self.extract_objects, 'target_folder')):
            self.external_link_paths = self.extract_objects.extract_file_paths()
        else:
            self.external_link_paths = {}

    def analyze(self):
        with open(self.path_for_test_file, 'w', encoding='utf-8') as file:
            for sheet_file in self.xml_content:
                split_sheet_dir = XMLSplitter.split_sheetdata(sheet_file)  # Получаем папку с данными для каждого листа
                xml_parser = SheetCellsParser(split_sheet_dir, self.external_link_paths)
                sys.stdout = file
                xml_parser.analyze_folder(self.extract_objects)  
                sys.stdout = sys.__stdout__
