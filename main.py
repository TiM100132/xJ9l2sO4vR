import os
import sys
import FileHandler
import XMLSplitter
import XMLParser

class Analyzer:
    def __init__(self, file_path):
        self.file_handler = FileHandler(file_path)
        self.xml_content = self.file_handler.extract_contents()  # Список XML файлов листов

    def analyze(self):
        with open(os.path.join(self.file_handler.temp_dir, "test_file.txt"), 'w', encoding='utf-8') as file:
            for sheet_file in self.xml_content:
                split_sheet_dir = XMLSplitter.split_sheetdata(sheet_file)  # Получаем папку с данными для каждого листа
                xml_parser = XMLParser(split_sheet_dir)
                sys.stdout = file
                xml_parser.analyze_folder()  
                sys.stdout = sys.__stdout__


if __name__ == "__main__":
    file_path = r""
    analyzer = Analyzer(file_path)
    analyzer.analyze()