import os
from lxml import etree



class ExtractObjectInfo:
    def __init__(self, temp_dir):
        self.temp_dir = temp_dir

        # vars for external_link
        self.target_folder = os.path.join(self.temp_dir, "xl", "externalLinks", "_rels")
        self.external_link_paths = {}

        # vars for workbook
        self.sheet_names = {}
        self.find_sheet_name()




    def extract_file_paths(self):
        try:
            for filename in os.listdir(self.target_folder):
                if filename.startswith("externalLink") and filename.endswith(".xml.rels"):
                    rel_id = filename[len("externalLink"):-len(".xml.rels")]
                    rel_file = os.path.join(self.target_folder, filename)
                    tree = etree.parse(rel_file)
                    ns = {"ns": "http://schemas.openxmlformats.org/package/2006/relationships"}
                    target_elements = tree.xpath("//ns:Relationship/@Target", namespaces=ns)
                    for target in target_elements:
                        self.external_link_paths[rel_id] = target

            return self.external_link_paths
        
        except Exception as e:
            print(f"Произошла ошибка: {e}")

    def find_sheet_name(self):
        workbook_path = os.path.join(self.temp_dir, "xl", "workbook.xml")
        try:
            tree = etree.parse(workbook_path)
            ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            target_elements = tree.xpath("//ns:sheet", namespaces=ns)
            for target in target_elements:
                self.sheet_names[target.attrib["sheetId"]] = target.attrib["name"]

        except Exception as e:
            print(f"Произошла ошибка: {e}")

    def get_sheet_name(self, sheet_id):
        return self.sheet_names.get(sheet_id)