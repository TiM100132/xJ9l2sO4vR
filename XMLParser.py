import os
import re
from lxml import etree
from multiprocessing import Pool
import RangeConvertor

class XMLParser:
    def __init__(self, sheet_folder):
        self.sheet_folder = sheet_folder
    
    def analize_xml(self, file):
        try:
            tree = etree.parse(file)
            root = tree.getroot()

            formula_cell_count = 0
            function_counts = {}
            numeric_cells = 0
            text_cells = 0
            empty_cells = 0

            for cell_element in root.iter('c'):
                formula_elements = cell_element.iter('f')
                for formula_element in formula_elements:
                    formula_cell_count += 1
                    if formula_element.text is not None:
                        formula_text = formula_element.text
                        function_name_matches = re.findall(r'([A-ZА-Я]+)\(', formula_text)
                        if "t" in formula_element.attrib and formula_element.attrib["t"] == "shared" and "ref" in formula_element.attrib:
                            for function_name in function_name_matches:
                                function_counts[function_name] = function_counts.get(function_name, 0) + RangeConvertor.count_cells_in_ranges(formula_element.get("ref"))
                        else:
                            for function_name in function_name_matches:
                                function_counts[function_name] = function_counts.get(function_name, 0) + 1

                value_element = cell_element.find('v')
                if value_element is not None or cell_element.text is not None:
                    cell_type = cell_element.get('t')
                    if cell_type == 'n':
                        numeric_cells += 1
                    elif cell_type == 's':
                        text_cells += 1
                    elif cell_type is None and value_element is None:
                        empty_cells += 1
                else:
                    empty_cells += 1

            return formula_cell_count, function_counts, numeric_cells, text_cells, empty_cells
        except Exception as e:
            print(f"Ошибка при анализе файла {file}: {e}")
            return 0, {}, 0, 0, 0
        
    def analyze_folder(self):
        
        if self.sheet_folder is not None:
            files = [os.path.join(root, file) for root, _, files in os.walk(self.sheet_folder) for file in files if file.endswith('.xml')]
            
            if not files:
                print("Папка не содержит XML файлов.")
            else:

                with Pool() as pool:
                    results = pool.map(self.analize_xml, files)

                total_formula_cell_count = 0
                total_function_counts = {}
                total_numeric_cells = 0
                total_text_cells = 0
                total_empty_cells = 0

                for formula_cell_count, function_counts, numeric_cells, text_cells, empty_cells in results:
                    total_formula_cell_count += formula_cell_count
                    for function_name, count in function_counts.items():
                        total_function_counts[function_name] = total_function_counts.get(function_name, 0) + count
                    total_numeric_cells += numeric_cells
                    total_text_cells += text_cells
                    total_empty_cells += empty_cells

                name_folder = os.path.basename(self.sheet_folder)
                print(name_folder)
                print(f"Количество ячеек с формулой в {name_folder}: {total_formula_cell_count}")
                print("Количество использований функций:")
                for function, count in total_function_counts.items():
                    print(f"{function}: {count} раз(а)")
                print(f"Количество числовых ячеек: {total_numeric_cells}")
                print(f"Количество текстовых ячеек: {total_text_cells}")
                print(f"Количество пустых ячеек: {total_empty_cells}")
                print('\n')