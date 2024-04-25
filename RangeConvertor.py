import re

class RangeConvertor:
    @staticmethod
    def count_cells_in_ranges(cell_ranges):
        """Рассчитывает общее количество ячеек в указанных диапазонах."""
        cell_ranges_list = cell_ranges.split()
        total_cells = 0

        for cell_range in cell_ranges_list:
            start_column, start_row, end_column, end_row = RangeConvertor.cell_range_to_indices(cell_range)
            num_cells = RangeConvertor.calculate_num_cells(start_column, end_column, start_row, end_row)
            total_cells += num_cells

        return total_cells
    
    @staticmethod
    def cell_range_to_indices(cell_range):
        """Преобразует буквенно-числовое обозначение ячейки в числовые индексы."""
        match = re.match(r'([A-Z]+)(\d+):?([A-Z]+)?(\d+)?', cell_range)
        start_column_str, start_row_str, end_column_str, end_row_str = match.groups()

        start_column = RangeConvertor.column_to_index(start_column_str)
        start_row = int(start_row_str)
        
        if end_column_str and end_row_str:
            end_column = RangeConvertor.column_to_index(end_column_str)
            end_row = int(end_row_str)
        else:
            # Если не указан конечный столбец и строка, считаем это за одну ячейку
            end_column = start_column
            end_row = start_row

        return start_column, start_row, end_column, end_row
    
    @staticmethod
    def column_to_index(column_str):
        """Преобразует буквенное обозначение столбца в числовой индекс."""
        column_str = column_str.upper()
        index = 0
        for char in column_str:
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index
    
    @staticmethod
    def calculate_num_cells(start_column, end_column, start_row, end_row):
        """Рассчитывает количество ячеек в заданном диапазоне."""
        num_columns = end_column - start_column + 1
        num_rows = end_row - start_row + 1
        return num_columns * num_rows