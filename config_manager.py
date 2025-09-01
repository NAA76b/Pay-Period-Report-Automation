import configparser

class ConfigManager:
    """Handles loading and saving application settings from config.ini."""
    def __init__(self, path='config.ini'):
        self.path = path
        self.config = configparser.ConfigParser()
        self.load_config()

    def load_config(self):
        """Loads the configuration from the file."""
        self.config.read(self.path)

    def get(self, section, key):
        """Gets a value from the config."""
        return self.config.get(section, key)

    def set(self, section, key, value):
        """Sets a value in the config."""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)

    def save_config(self):
        """Saves the current configuration to the file."""
        with open(self.path, 'w') as configfile:
            self.config.write(configfile)
