import os
import re
from lxml import etree
from multiprocessing import Pool
from range_convertor import RangeConvertor
from extract_ooxml_object_info import ExtractObjectInfo

class SheetCellsParser:
    def __init__(self, sheet_folder, external_link_paths):
        self.sheet_folder = sheet_folder
        self.formula_cell_count = 0
        self.function_counts = {}
        self.numeric_cells = 0
        self.text_cells = 0
        self.empty_cells = 0
        
        # vars for conditional_formattings
        self.total_cf_count = 0
        self.cf_type_counts = {}
        self.formatted_data = []

        # vars for external_link
        self.external_link_paths = external_link_paths
        self.updated_result_for_sheet = {}
        self.total_count_external_link = 0  

    
    def analyze_xml(self, file):
        try:
            tree = etree.parse(file)
            root = tree.getroot()

            for cell_element in root.iter('c'):
                formula_elements = cell_element.iter('f')
                for formula_element in formula_elements:
                    self.formula_cell_count += 1
                    if formula_element.text is not None:
                        formula_text = formula_element.text

                        if self.external_link_paths:
                            total_count, updated_results = self.count_file_paths_in_formulas(formula_text)
                            self.total_count_external_link += total_count
                            self.updated_result_for_sheet.update(updated_results)
                        else:
                            total_count, updated_results = 0, {}
                            
                        function_name_matches = re.findall(r'([A-ZА-Я]+)\(', formula_text)
                        if "t" in formula_element.attrib and formula_element.attrib["t"] == "shared" and "ref" in formula_element.attrib:
                            for function_name in function_name_matches:
                                self.function_counts[function_name] = self.function_counts.get(function_name, 0) + RangeConvertor.count_cells_in_ranges(formula_element.get("ref"))
                        else:
                            for function_name in function_name_matches:
                                self.function_counts[function_name] = self.function_counts.get(function_name, 0) + 1

                value_element = cell_element.find('v')
                if value_element is not None or cell_element.text is not None:
                    cell_type = cell_element.get('t')
                    if cell_type == 'n':
                        self.numeric_cells += 1
                    elif cell_type == 's':
                        self.text_cells += 1
                    elif cell_type is None and value_element is None:
                        self.empty_cells += 1
                else:
                    self.empty_cells += 1
            # print(self.updated_result_for_sheet)
            return self.formula_cell_count, self.function_counts, self.numeric_cells, self.text_cells, self.empty_cells, self.total_count_external_link, self.updated_result_for_sheet
        except Exception as e:
            print(f"Ошибка при анализе файла {file}: {e}")
            return 0, {}, 0, 0, 0
        

    def count_file_paths_in_formulas(self, string):
        matches = re.findall(r"\[(\d+)]", string)
        total_count_external_link = 0
        for match in matches:
            total_count_external_link += 1  
            if match in self.external_link_paths:
                if self.external_link_paths[match] not in self.updated_result_for_sheet:
                    self.updated_result_for_sheet[self.external_link_paths[match]] = 0
                self.updated_result_for_sheet[self.external_link_paths[match]] += 1
            else:
                if match not in self.updated_result_for_sheet:
                    self.updated_result_for_sheet[match] = 0
                self.updated_result_for_sheet[match] += 1

        return total_count_external_link, self.updated_result_for_sheet


    def analyze_conditional_formatting(self, file):
        try:
            tree = etree.parse(file)
            root = tree.getroot()

            for cell_element in root.iter('conditionalFormatting'):
                self.total_cf_count += 1
                sqref = cell_element.get('sqref')
                cf_rule_element = cell_element.find('cfRule')
                rule_type = cf_rule_element.attrib["type"]
                self.cf_type_counts[rule_type] = self.cf_type_counts.get(rule_type, 0) + 1

                self.formatted_data.append({
                        'sqref': sqref,
                        'rule_type': rule_type
                    })
        
        except Exception as e:
                print(f"Ошибка при анализе файла {file}: {e}")

    def analyze_folder(self, extract_objects):
        
        if self.sheet_folder is not None:

            total_formula_cell_count = 0
            total_function_counts = {}
            total_numeric_cells = 0
            total_text_cells = 0
            total_empty_cells = 0

            
            files = [os.path.join(root, file) for root, _, files in os.walk(self.sheet_folder) for file in files if file.endswith('.xml')]
            conditional_formatting_dir = os.path.join(self.sheet_folder, 'conditionalFormatting')
            
            if os.path.exists(conditional_formatting_dir) and os.path.isdir(conditional_formatting_dir):
                files_in_dir = os.listdir(conditional_formatting_dir)
                file_path = os.path.join(conditional_formatting_dir, files_in_dir[0])
                self.analyze_conditional_formatting(file_path)

            
            if not files:
                print("Папка не содержит XML файлов.")
            else:

                with Pool() as pool:
                    results = pool.map(self.analyze_xml, files)

                

                for formula_cell_count, function_counts, numeric_cells, text_cells, empty_cells, total_count, updated_results in results:
                    total_formula_cell_count += formula_cell_count
                    for function_name, count in function_counts.items():
                        total_function_counts[function_name] = total_function_counts.get(function_name, 0) + count
                    total_numeric_cells += numeric_cells
                    total_text_cells += text_cells
                    total_empty_cells += empty_cells
                    self.total_count_external_link += total_count
                    self.updated_result_for_sheet.update(updated_results)


                

                sheet_id = os.path.basename(self.sheet_folder)[5:]
                name_sheet = extract_objects.get_sheet_name(sheet_id)

                print(f'Лист - {name_sheet}')
                print(f"Количество ячеек с формулой на листе {name_sheet}: {total_formula_cell_count}")
                print("Количество использований функций:")
                for function, count in total_function_counts.items():
                    print(f"{function}: {count} раз(а)")
                print(f"Количество числовых ячеек: {total_numeric_cells}")
                print(f"Количество текстовых ячеек: {total_text_cells}")
                print(f"Количество пустых ячеек: {total_empty_cells}")
                print(f"Количество внешних ссылок: {self.total_count_external_link}")
                
                print(f"Общее количество правил условного форматирований: {self.total_cf_count}")
                if self.total_cf_count:
                    print("Распределение по типу форматирования:")
                    for cf_type, count in self.cf_type_counts.items():
                        print(f"Тип: {cf_type}, Количество: {count}")

                    for data in self.formatted_data:
                        print(f"Диапазон: {data['sqref']}, Тип форматирования: {data['rule_type']}")


                if self.total_count_external_link != 0:
                    print("___________Список файлов на которые добавлены внешние ссылки___________")
                    for index, count in self.updated_result_for_sheet.items():
                        print(f"Ссылка на документ: {index}, Количество ячеек: {count}")

                print('\n')


