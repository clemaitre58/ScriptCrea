import csv
import os
class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file = None):
        self._list_project = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        self._data_path = self._dir_proj + '/Data'
        
        if (self._existing_out_file == None):
            self._existing_out_file = self._dir_proj + '/Suivi/suivi.csv'
            self._create_empty_ofile(self._existing_out_file)


    # def check_fiche_vs_xls():

    def _create_empty_ofile(self, name):
        with open(name, 'wb') as fl:
            fl.close()

    def _create_new_trimester(self, name_tri, name_year):

        if (name_tri == 'T1' 
                or name_tri == 'T2'
                or name_tri == 'T3'
                or name_tri == 'T4'):
            name_dir_to_list = self._data_path + '/' + name_year + '/' + name_tri + '/'
            self._all_tri_fiches = self._list_fiche_directory(name_dir_to_list)
            # load base_mission
        else:
            print "Given value for name au trimester not allow! Waiting T1, T2, T3 or T4"

    def _list_fiche_directory(sefl, dir_path):
        list_files = []
        for root, dirs, files in os.walk(dir_path): 
            for nom_fichier in files:
                list_files.append(os.path.join(root, nom_fichier))

        for idx, iten in enumerate(list_files):
            if( list_files[idx].find("base_mission") != -1):
                del list_files[idx]
        
        return list_files

    def print_all_tri_fiches(self):
        for name in self._all_tri_fiches:
            print name
