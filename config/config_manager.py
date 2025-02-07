import configparser
import os
import re

def load_config(file_path):
    config = configparser.ConfigParser()
    config.optionxform = str  # preserve case for keys
    config.read(file_path)
    result = {}
    for section in config.sections():
        result[section] = dict(config[section])
    return result

def save_config(file_path, config_dict):
    config = configparser.ConfigParser()
    config.optionxform = str  # preserve case for keys
    for section, values in config_dict.items():
        config[section] = {}
        for key, value in values.items():
            config[section][key] = str(value)
    with open(file_path, "w") as f:
        config.write(f)

def default_config():
    return {
        "GENERAL": {
            "TITLE": "The Series S1-S3 Production Calendar",
            "COLUMN_WIDTH": "4.9"
        },
        "PHASE_COLORS": {
            "Development": "FFFF00",
            "Pre-pre-production": "FFA500",
            "Pre-production": "83F28F",
            "Shooting": "00C04B",
            "Post production": "7C4700",
            "Financing": "737CA1",
            "Marketing": "82CAFF",
            "Premier": "FFC0CB"
        },
        "ROW_HEIGHTS": {
            "normal": "20",
            "special": "60"
        }
    }

def sanitize_title(title):
    # Remove characters that are not alphanumeric, hyphen, underscore, or space.
    sanitized = re.sub(r"[^\w\s-]", "", title).strip()
    sanitized = re.sub(r"[\s]+", "_", sanitized)
    return sanitized

def save_project_config(config_dict, folder):
    title = config_dict.get("GENERAL", {}).get("TITLE", "default_project")
    if not title.strip():
        title = "default_project"
    filename = sanitize_title(title) + ".ini"
    file_path = os.path.join(folder, filename)
    save_config(file_path, config_dict)
    return file_path