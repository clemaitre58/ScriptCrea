from utils.etude import *

EtudeCrea2016 = EtudeFiches('/home/cedric/Documents/Conseil/Creative',None)
EtudeCrea2016._create_new_trimester('T4', '2016')
#EtudeCrea2016.print_all_tri_fiches()
EtudeCrea2016.match_fiche_data()
EtudeCrea2016.show_projects_fiches_missing()
#EtudeCrea2016.show_all_project_data()
EtudeCrea2016.xlxs_to_csv('/home/cedric/Documents/Conseil/Creative/Suivi/Matrice_suivi_fiche.xlsx', None)
EtudeCrea2016.proc_trimester('3')
#EtudeCrea2016.proc_proj_a_traiter()
EtudeCrea2016.export_l_csv('/home/cedric/Documents/Conseil/Creative/Suivi/Matrice_suivi_fiche.csv', EtudeCrea2016._liste_a_traiter)
