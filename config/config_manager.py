import configparser

def load_config(file_path):
    config = configparser.ConfigParser()
    config.read(file_path)
    # Convert to a nested dictionary
    result = {}
    for section in config.sections():
        result[section] = dict(config[section])
    return result

def save_config(file_path, config_dict):
    config = configparser.ConfigParser()
    for section, values in config_dict.items():
        config[section] = {}
        for key, value in values.items():
            config[section][key] = str(value)
    with open(file_path, "w") as f:
        config.write(f)

def default_config():
    return {
        "GENERAL": {
            "TITLE": 'The Series S1-S3 Production Calendar',
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