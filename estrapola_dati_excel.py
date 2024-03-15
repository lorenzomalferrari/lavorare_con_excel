#!/usr/bin/env python3

import pandas as pd
import json

import re
import logging
import pathlib
import argparse
import psycopg2
import psycopg2.extras
import coloredlogs

class ExcelImporter:
    '''
        Dato un file excel, estrapolare dati
        e metterli in un:
        - dizionario
        - json
    '''

    def __init__(self, excel_path_str:str) -> None:
        self.log = logging.getLogger('excel')
        excel_path = pathlib.Path(excel_path_str)

        self.log.debug('Leggo il contenuto del file %s', excel_path)

        # Carica il file Excel
        df = pd.read_excel(excel_path, sheet_name=None)

        # Crea un dizionario per ciascun foglio del file Excel
        data_dict = {}
        for sheet_name, data_frame in df.items():
            # Converti Timestamp in stringhe
            data_frame = data_frame.applymap(lambda x: str(x) if isinstance(x, pd.Timestamp) else x)
            data_dict[sheet_name] = data_frame.to_dict(orient='records')

        # Salva il dizionario in un file JSON
        with open('dati.json', 'w') as json_file:
            json.dump(data_dict, json_file, indent=4)

        print("Il file JSON è stato creato con successo.")

if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument('-v', '--verbose', action='store_true', help='Show more output')
	parser.add_argument('-q', '--quiet', action='store_true', help='Show less output')
	parser.add_argument('excelfile', help='Excel file')
	args = parser.parse_args()

	if args.verbose:
		ll = logging.DEBUG
	elif args.quiet:
		ll = logging.WARNING
	else:
		ll = logging.INFO
	logging.basicConfig(level=ll)
	coloredlogs.install(level=ll)

	ExcelImporter(args.excelfile)
