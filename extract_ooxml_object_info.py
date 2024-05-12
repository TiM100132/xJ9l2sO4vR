import os
from lxml import etree



class ExtractObjectInfo:
    def __init__(self, temp_dir):
        self.temp_dir = temp_dir

        # vars for external_link
        self.target_folder = os.path.join(self.temp_dir, "xl", "externalLinks", "_rels")
        self.external_link_paths = {}
        # if os.path.exists(self.target_folder):
        #     self.extract_file_paths()


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