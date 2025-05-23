## utility functions 
import os
import random
import string
import datetime
import json
import os

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")
TEMPLATE_CONFIGS_FOLDER = os.path.join(os.path.dirname(__file__), "..", "template_configs")

def load_global_settings():
    """Load global app settings from 'settings.json'."""
    if not os.path.exists(SETTINGS_FILE):
        # Return default settings if the file doesn't exist
        return {
            "active_template": "Default",
            "delimiter_extract": "|",
            "delimiter_zip": "|",
            "enable_header": True,
            "enable_trailer": True
        }

    try:
        with open(SETTINGS_FILE, "r") as file:
            return json.load(file)
    except json.JSONDecodeError:
        raise ValueError("Failed to parse 'settings.json'. Please check file formatting.")

def load_template_settings(template_name):
    """Load template-specific settings from JSON configuration files."""
    template_file = os.path.join(TEMPLATE_CONFIGS_FOLDER, f"{template_name}.json")

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"Template configuration '{template_name}.json' not found in {TEMPLATE_CONFIGS_FOLDER}")

    try:
        with open(template_file, "r") as file:
            return json.load(file)
    except json.JSONDecodeError:
        raise ValueError(f"Failed to parse '{template_name}.json'. Please check file formatting.")


def generate_random_value(pattern):
    """Generates a random value based on a pattern definition."""
    if pattern == "{random_6_digits}":
        return str(random.randint(100000, 999999))
    elif pattern == "{random_5_alphanumeric}":
        return ''.join(random.choices(string.ascii_letters + string.digits, k=5))
    elif pattern == "{random_L5N}":
        return random.choice(string.ascii_uppercase) + str(random.randint(10000, 99999))
    return pattern  # Return static text if it's not a pattern

def replace_placeholders(template, file_name, row_count):
    """Replaces placeholders in the header/trailer format string."""
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    template = template.replace("{timestamp}", timestamp)
    template = template.replace("{row_count}", str(row_count))
    template = template.replace("{file_name}", file_name)

    # Replace random value placeholders
    for pattern in ["{random_6_digits}", "{random_5_alphanumeric}", "{random_L5N}"]:
        while pattern in template:
            template = template.replace(pattern, generate_random_value(pattern), 1)

    return template


##prevent overwriting if filename already ecists and instantiate a versioned name
def get_unique_filename(filepath):
    base, ext = os.path.splitext(filepath)
    counter = 1
    new_filepath = filepath
    while os.path.exists(new_filepath):
        new_filepath = f"{base} ({counter}){ext}"
        counter += 1
    return new_filepath
