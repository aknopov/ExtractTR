import configparser
import os


class Prefs:
    def __init__(self):
        self.os = os
        self.config = configparser.ConfigParser()
        self.prefs_path = os.path.join(os.path.expanduser("~"), '.extractrc')
        self.dirs = ()

    def get_dirs(self):
        if self.dirs == tuple():
            self.config.read(self.prefs_path)
            unnamed = self.config[self.config.default_section]
            src_dir = unnamed.get('source', '.')
            dest_dir = unnamed.get('destination', '.')
            self.dirs = (src_dir, dest_dir)
        return self.dirs

    def save_dirs(self, src_dir, dest_dir):
        self.dirs = (src_dir, dest_dir)
        self.config[self.config.default_section] = {
            'source': src_dir,
            'destination': dest_dir
        }
        with open(file=self.prefs_path, mode='w', encoding='UTF-8') as configfile:
            self.config.write(configfile)
