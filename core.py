import pickle

from argparse import ArgumentParser
from pathlib import Path

import excel

def load_template(template, func):
    try:
        with open(template, 'rb') as file:
            template = pickle.load(file)[func]
    except FileNotFoundError:
        return False
    else:
        return {'temp_sheet': template['template'], 'paths': template['paths']}



def miso(files, output, market, template_path):
    template = load_template(template_path, market)
    rotater = excel.HandlerRotater()