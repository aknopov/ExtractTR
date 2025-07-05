import configparser
import os

prefs_path = os.path.join(os.path.expanduser("~"), '.extractrc')
config = configparser.ConfigParser()
dirs = ()

def get_dirs():
    global dirs
    if dirs == tuple():
        config.read(prefs_path)
        unnamed = config[config.default_section]
        src_dir = unnamed.get('source', '.')
        dest_dir = unnamed.get('destination', '.')
        dirs = (src_dir, dest_dir)
    return dirs

def save_dirs(src_dir, dest_dir):
    global dirs
    dirs = (src_dir, dest_dir)
    config[config.default_section] = {
        'source': src_dir,
        'destination': dest_dir
    }
    with open(file=prefs_path, mode='w', encoding='UTF-8') as configfile:
        config.write(configfile)
