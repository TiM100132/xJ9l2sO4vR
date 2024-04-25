import os
import re

class XMLSplitter:
    @staticmethod
    def split_sheetdata(input_file):
        with open(input_file, 'r', encoding='utf-8') as file:
            xml_data = file.read()

        output_directory = os.path.join(os.path.dirname(input_file), 'analizing_data', os.path.splitext(os.path.basename(input_file))[0])
        os.makedirs(output_directory, exist_ok=True)

        # Блок для разделения элементов в <conditionalFormatting>
        pattern_for_cF = re.compile(r'<conditionalFormatting.*?</conditionalFormatting>', re.DOTALL)
        matches_cF = pattern_for_cF.findall(xml_data)
        
        if matches_cF:
            cF_dir = os.path.join(os.path.dirname(input_file), 'analizing_data', 'conditionalFormatting')
            os.makedirs(cF_dir, exist_ok=True)
            with open(os.path.join(cF_dir, f'{os.path.splitext(os.path.basename(input_file))[0]}.xml'), 'w', encoding='utf-8') as f:
                for match in matches_cF:
                    f.write(match + '\n')

        # Блок для разделения элементов в <sheetData>
        dimension_pattern = re.compile(r'<dimension ref="([^"]+)"')
        match = dimension_pattern.search(xml_data)
        last_number = None
        if match:
            ref_string = match.group(1)
            numbers = re.findall(r'\d+', ref_string)
            if numbers:
                last_number = int(numbers[-1])

        if last_number is not None:
            if last_number < 100000:
                row_limit = 10000
            else:
                row_limit = last_number // 10

            sheetdata_pattern = re.compile(r'<sheetData>.*?</sheetData>', re.DOTALL)
            sheetdata_match = sheetdata_pattern.search(xml_data)
            if not sheetdata_match:
                return
            del xml_data
            sheetdata = sheetdata_match.group()

            row_pattern = re.compile(r'<row[^>]*>.*?</row>', re.DOTALL)
            rows = row_pattern.findall(sheetdata)
            del sheetdata
            row_limit = 1000
            row_groups = [rows[i:i + row_limit] for i in range(0, len(rows), row_limit)]
            del rows

            for i, group in enumerate(row_groups, start=1):
                with open(os.path.join(output_directory, f'sheetData_part_{i}.xml'), 'w', encoding='utf-8') as part_file:
                    part_file.write('<sheetData>\n')
                    for row in group:
                        part_file.write(row + '\n')
                    part_file.write('</sheetData>')
            
            return output_directory

        else:
            print("Не удалось найти информацию о размере данных в файле.")