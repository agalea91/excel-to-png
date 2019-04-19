# -*- coding: utf-8 -*-
import excel2img
import glob
import click
import os
from tqdm import tqdm

DATA_PATH = '../xlsx'

@click.command()
@click.option(
    '--target-file',
    default=None,
    help='Filetype'
)
@click.option(
    '--sheet-name',
    default='',
    help=''
)
@click.option(
    '--all-cells',
    is_flag=True
)
@click.option(
    '--cell-range',
    default=None
)
class Main():
    def __init__(self, target_file, sheet_name, all_cells, cell_range):
        self.target_file = target_file
        self.sheet_name = sheet_name
        self.all_cells = all_cells
        self.cell_range = cell_range

    def run(self):
        if not self.target_file:
            target_files = glob.glob(os.path.join(DATA_PATH, '*'))
        else:
            target_files = [self.target_file]
        
        for file in tqdm(target_files):
            self.convert_to_png(file)


    def convert_to_png(self, file):
        args = [
            file,
            '{}.png'.format(os.path.splitext(file)[0]),
            '',
            ''
        ]
        if self.all_cells:
            args[2] = self.sheet_name
        if cell_range:
            args[3] = '{}!{}'.format(self.sheet_name, self.cell_range)

        try:
            excel2img.export_img(*args)
        except Exception as e:
            if self.raise_errors:
                raise(e)
            print('Exception raised in excel2img.export_img.')
            print(e)
            print('Image not saved')


if __name__ == '__main__':
    Main.run()