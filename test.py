from utils.etude import EtudeFiches

EtudeCrea2016 = EtudeFiches('/home/cedric/Documents/Conseil/Creative', None)
EtudeCrea2016._create_new_trimester('T4', '2016')
# EtudeCrea2016.print_all_tri_fiches()
EtudeCrea2016.match_fiche_data()
EtudeCrea2016.show_projects_fiches_missing()
# EtudeCrea2016.show_all_project_data()
EtudeCrea2016.xlxs_to_csv(
    '/home/cedric/Documents/Conseil/Creative/Suivi/Matrice_suivi_fiche.xlsxr',
    None)
EtudeCrea2016.proc_trimester('3')
EtudeCrea2016.clean_l_proj(EtudeCrea2016._list_project)
EtudeCrea2016.proc_proj_a_traiter()
f_exp_at = ("/home/cedric/Documents/Conseil/Creative/Suivi/"
            "Matrice_suivi_fiche_at.html")
f_exp_l_html = ("/home/cedric/Documents/Conseil/Creative/Suivi/"
                "suivi_fiche_2016.html")
f_exp_l_csv = ("/home/cedric/Documents/Conseil/Creative/Suivi/"
               "suivi_fiche_2016.csv")
EtudeCrea2016.export_l_html(f_exp_at, EtudeCrea2016._liste_a_traiter)
EtudeCrea2016.export_l_html(f_exp_l_html, EtudeCrea2016._list_project)
EtudeCrea2016.export_l_csv(f_exp_l_html, EtudeCrea2016._list_project)
