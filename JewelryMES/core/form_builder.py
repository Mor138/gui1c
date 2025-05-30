# JewelryMES/core/form_builder.py

import os
import json
import xml.etree.ElementTree as ET
from pathlib import Path

SCHEMA_DIR = Path("form_schemas")
SCHEMA_DIR.mkdir(exist_ok=True)

FIELD_TYPES = ["Строка", "Число", "Дата", "Булево"]

class FormSchemaManager:
    def __init__(self, schema_dir=SCHEMA_DIR):
        self.schema_dir = schema_dir

    def load_schema(self, stage_key):
        path = self.schema_dir / f"{stage_key}.json"
        if not path.exists():
            return []
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_schema(self, stage_key, fields):
        path = self.schema_dir / f"{stage_key}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(fields, f, ensure_ascii=False, indent=2)

    def add_field(self, stage_key, name, dtype):
        fields = self.load_schema(stage_key)
        fields.append({"name": name, "type": dtype})
        self.save_schema(stage_key, fields)

    def list_schemas(self):
        return [p.stem for p in self.schema_dir.glob("*.json")]


class FormParser:
    def __init__(self, base_path="conf"):
        self.base_path = Path(base_path)

    def parse_all_forms(self):
        fields_by_form = {}
        for root, _, files in os.walk(self.base_path):
            for file in files:
                if file == "Form.xml":
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, self.base_path)
                    try:
                        tree = ET.parse(full_path)
                        form_root = tree.getroot()
                        fields = []
                        for field in form_root.findall(".//{*}LabelField"):
                            fields.append({
                                "name": field.attrib.get("name", "—"),
                                "path": field.findtext("{*}DataPath", default="")
                            })
                        fields_by_form[rel_path] = fields
                    except Exception as e:
                        fields_by_form[rel_path] = [{"error": str(e)}]
        return fields_by_form

    def export_to_json(self, output_file="structured_form_fields.json"):
        data = self.parse_all_forms()
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
