from utils.etude import EtudeFiches

EtudeCrea2017 = EtudeFiches('/home/cedric/Documents/Conseil/Creative',
                            '/Suivi/suivi_fiche_2016.csv')

p_data_2017 = '/home/cedric/Documents/Conseil/Creative/Suivi/Data_2017.csv'

l_continu, l_close_p,  l_new_proj = EtudeCrea2017.import_new_year(p_data_2017)

l_merge = l_continu + l_close_p + l_new_proj

l_update = EtudeCrea2017.process_from_t(l_merge, 3)

# EtudeCrea2016._create_new_trimester('T4', '2016')
# # EtudeCrea2016.print_all_tri_fiches()
# EtudeCrea2016.match_fiche_data()
# EtudeCrea2016.show_projects_fiches_missing()
# # EtudeCrea2016.show_all_project_data()
#
# source_suivi = \
#     '/home/cedric/Documents/Conseil/Creative/Suivi/Matrice_suivi_fiche.xlsx'
#
# EtudeCrea2016.xlxs_to_csv(source_suivi, None)
# EtudeCrea2016.proc_trimester('3')
# EtudeCrea2016.clean_l_proj(EtudeCrea2016._list_project)
# EtudeCrea2016.proc_proj_a_traiter()
# f_suivi_fiche_at_html = \
#     '/home/cedric/Documents/Conseil/Creative/Suivi/Matrice_suivi_fiche_at.html'
# EtudeCrea2016.export_l_html(f_suivi_fiche_at_html,
#                             EtudeCrea2016._liste_a_traiter)
# f_suivi_fiche_2016_html = \
#     '/home/cedric/Documents/Conseil/Creative/Suivi/suivi_fiche_2016.html'
# EtudeCrea2016.export_l_html(f_suivi_fiche_2016_html,
#                             EtudeCrea2016._list_project)
# f_suivi_fiche_2016_csv = \
#     '/home/cedric/Documents/Conseil/Creative/Suivi/suivi_fiche_2016.csv'
# EtudeCrea2016.export_l_csv(
# f_suivi_fiche_2016_csv, EtudeCrea2016._list_project)
