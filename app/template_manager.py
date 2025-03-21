import json
import os
import shutil
from datetime import datetime
import random

# Define paths for config file and templates folder
CONFIG_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "config.json"))
TEMPLATES_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "templates"))


class TemplateManager:
    """Manages template configurations and Excel template uploads."""

    def __init__(self):
        self.templates = self.load_templates()

    def load_templates(self):
        """Load templates from the config file or initialize with default template."""
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, 'r') as file:
                    return json.load(file).get('templates', {})
            except json.JSONDecodeError as e:
                print(f"Error loading config.json: {e}")
                return {"TaxOverlayS3Data": self.default_template()}
        return {"TaxOverlayS3Data": self.default_template()}


    def save_templates(self):
        with open(CONFIG_PATH, 'w') as file:
            json.dump({"templates": self.templates}, file, indent=4)

    def default_template(self):
        return {
            "delimiter": "|",
            "header_records": 1,
            "trailer_records": 1,
            "zip_naming_convention": "TAXOPT.LPB.{timestamp}.{trading_request_id}",
            "header_format": "HDR|LPB|{trading_request_id}|{timestamp}|{sheet_name}",
            "trailer_format": "TLR|{row_count}",
            "txt_filename_pattern": "TAXOPT.{timestamp}.{trading_request_id}.{sheet_name}.txt"
        }

    def get_template(self, template_name='TaxOverlayS3Data'):
        return self.templates.get(template_name, self.default_template())

    def resolve_tokens(self, text, **kwargs):
        """Dynamically replace tokens like {timestamp}, {trading_request_id}, etc."""
        tokens = {
            "timestamp": datetime.datetime.now.strftime("%Y%m%d:%H:%M:%S:%f")[:-3],
            "trading_request_id": random.randint(100000, 999999),
            **kwargs  # Additional dynamic data (e.g., `sheet_name`, `row_count`)
        }
        try:
            return text.format(**tokens)
        except KeyError as e:
            raise ValueError(f"Missing token in template: {e}")


    def create_template(self, template_name, delimiter='|', header_records=0, trailer_records=0):
        """Create or update a template dynamically from user entries."""
        self.templates[template_name] = {
            "delimiter": delimiter,
            "header_records": header_records,
            "trailer_records": trailer_records,
            "zip_naming_convention": f"{template_name}_ZIP_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "header_format": "HDR|{trading_request_id}|{timestamp}|{sheet_name}",
            "trailer_format": "TLR|{row_count}",
            "txt_filename_pattern": f"{template_name}_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
        }
        self.save_templates()

    def delete_template(self, template_name):
        if template_name in self.templates:
            del self.templates[template_name]
            self.save_templates()
            return True
        return False
    
    def list_templates(self):
    # """Return a list of available template names."""
        return list(self.templates.keys())

