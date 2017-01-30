import csv

class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file = None):
        self._list_project = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        
        if (self._existing_out_file == None):
            self._existing_out_file = self._dir_proj + '/Suivi/suivi.csv'
            _create_empty_ofile(self._existing_out_file)


    def check_fiche_vs_xls():

    def _create_empty_ofile(self, name):
        myfile = open(name, 'wb')
        myfile.close()




