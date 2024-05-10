import os
import sys
from file_handler import FileHandler
from xml_splitter import XMLSplitter
from xml_parser import XMLParser

class Analyzer:
    def __init__(self, file_path):
        self.file_handler = FileHandler(file_path)
        self.xml_content = self.file_handler.extract_contents()  # Список XML файлов листов
        self.path_for_test_file = os.path.join(os.path.dirname(file_path), "test_file.txt")

    def analyze(self):
        with open(self.path_for_test_file, 'w', encoding='utf-8') as file:
            for sheet_file in self.xml_content:
                split_sheet_dir = XMLSplitter.split_sheetdata(sheet_file)  # Получаем папку с данными для каждого листа
                xml_parser = XMLParser(split_sheet_dir)
                sys.stdout = file
                xml_parser.analyze_folder()  
                sys.stdout = sys.__stdout__
