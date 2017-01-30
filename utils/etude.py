import csv

class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file = None):
        self._list_project = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        
        if (self._existing_out_file == None):
            self._existing_out_file = self._dir_proj + '/Suivis/suivi.csv'
            self._create_empty_ofile(self._existing_out_file)


    # def check_fiche_vs_xls():

    def _create_empty_ofile(self, name):
        with open(name, 'wb') myfile :
        myfile.close()

    def _create_new_trimester(self, name_tri):

        if name_tri == 'T1':
        
        elif name_tri == 'T2':

        elif name_tri == 'T3':

        elif name_tri == 'T4':

        else:
            print "Given value for name au trimester not allow \
                    waiting T1, T2, T3 or T4"





