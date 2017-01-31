from utils.etude import *

EtudeCrea2016 = EtudeFiches('/home/cedric/Documents/Conseil/Creative',None)
EtudeCrea2016._create_new_trimester('T4', '2016')
#EtudeCrea2016.print_all_tri_fiches()
EtudeCrea2016.match_fiche_data()
EtudeCrea2016.show_projects_fiches_missing()
#EtudeCrea2016.show_all_project_data()
