# import csv
from copy import deepcopy
import os
import xlrd
import subprocess
from utils.project import Project
from utils.str import similarity
import pandas as pd


class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file=None):
        self._list_project = []
        self._list_project_csv = []
        self._list_project_xlsx = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        self._data_path = self._dir_proj + '/Data'
        self._l_project_without_fiche = []
        self._list_nouveau = []
        self._list_conti = []
        self._list_petite_fiche = []
        self._liste_a_traiter = []
        self._liste_fiche = []
        self._liste_fiche_full = []
        self._l_new_year = []
        self._Debug = True
        if (self._existing_out_file is None):
            self._existing_out_file = self._dir_proj + '/Suivi/suivi.csv'
            self._create_empty_ofile(self._existing_out_file)
        else:
            if(self._Debug is True):
                print("Nom de fichier de données existant : ",
                      self._existing_out_file)
            self._existing_out_file = self._dir_proj + self._existing_out_file
            self._import_existing_suivi()
            print("importation " + self._existing_out_file + " Done!")

    def _create_empty_ofile(self, name):
        with open(name, 'wb') as fl:
            fl.close()

    def _create_new_trimester(self, name_tri, name_year):

        if (name_tri == 'T1'
                or name_tri == 'T2'
                or name_tri == 'T3'
                or name_tri == 'T4'):
            name_dir_to_list = self._data_path
            +'/' + name_year + '/' + name_tri + '/'
            self._all_tri_fiches = self._list_fiche_directory(name_dir_to_list)
            # load base_mission
            self._parse_xls_base_mission(name_dir_to_list, name_tri)
        else:
            print("Given value for name au trimester not allow!"
                  "Waiting T1, T2, T3 or T4")

    def _list_fiche_directory(sefl, dir_path):
        list_files = []
        for root, dirs, files in os.walk(dir_path):
            for nom_fichier in files:
                list_files.append(os.path.join(root, nom_fichier))

        for idx, iten in enumerate(list_files):
            if(list_files[idx].find("base_mission") != -1):
                del list_files[idx]

        return list_files

    def print_all_tri_fiches(self):
        for name in self._all_tri_fiches:
            print(name)

    def _parse_xls_base_mission(self, dir_path, name_tri):
        i_book_atl = -1
        i_book_ing = -1
        name_base_mission = dir_path + "base_mission.xlsx"
        book = xlrd.open_workbook(name_base_mission)
        print("Base mission : OK")  # + name_base_mission
        names_sheet = book.sheet_names()

        for idx, item in enumerate(names_sheet):
            if(names_sheet[idx].find("Fiche nombre de jour") != -1):
                i_book_ing = idx

        if(i_book_atl == -1 and i_book_ing == -1):
            print("Impossible de retourner les onglet dans le fichie excel")

        l_idx = i_book_ing
        for idx_b, item_b in enumerate(l_idx):

            all_data = book.sheet_by_index(item_b)
            l_titre = all_data.col_values(0)
            # ajout lect col nouv proj
            # Attention revoir tous les numéros
            l_is_new = all_data.col_values(6)
            if name_tri == "T1":
                l_T1 = all_data.col_values(1)
                l_T2 = None
                l_T3 = None
                l_T4 = None
            elif name_tri == "T2":
                l_T1 = all_data.col_values(1)
                l_T2 = all_data.col_values(2)
                l_T3 = None
                l_T4 = None
            elif name_tri == "T3":
                l_T1 = all_data.col_values(1)
                l_T2 = all_data.col_values(2)
                l_T3 = all_data.col_values(3)
                l_T4 = None
            elif name_tri == "T4":
                l_T1 = all_data.col_values(1)
                l_T2 = all_data.col_values(2)
                l_T3 = all_data.col_values(3)
                l_T4 = all_data.col_values(4)
            else:
                print("Valeurs pour name_tri mauvaise")

            # necessiterat un traitement a l'avenir
            # quand T2, T3, T4 n'existe pas

            for idx, item in enumerate(l_titre):
                if name_tri == "T1":
                    l_T = [l_T1[idx], None, None, None]
                elif name_tri == "T2":
                    l_T = [l_T1[idx], l_T2[idx], None, None]
                elif name_tri == "T3":
                    l_T = [l_T1[idx], l_T2[idx], l_T3[idx], None]
                elif name_tri == "T4":
                    l_T = [l_T1[idx], l_T2[idx], l_T3[idx], l_T4[idx]]
                else:
                    print("Valeurs pour name_tri mauvaise")
                print(l_titre[idx] + " " + str(l_T[0]) + " " + str(l_T[1]) +
                      " " + str(l_T[2]) + " " + str(l_T[3]))
                if l_is_new[idx] == "x":
                    is_new = True
                else:
                    is_new = False
                self._list_project_xlsx.append(Project(l_titre[idx], l_T,
                                                       None, is_new))

    def match_fiche_data(self):
        b_match = False
        if(self._list_project_xlsx is not None
                and self._all_tri_fiches is not None):
            for idx_p, item_p in enumerate(self._list_project_xlsx):
                b_match = False
                for idx_f, item_f in enumerate(self._all_tri_fiches):
                    if(self._all_tri_fiches[idx_f].find(
                            self._list_project_xlsx[idx_p]._name_project != -1)):
                        self._list_project_xlsx[idx_p].add_fiche(
                            self._all_tri_fiches[idx_f])
                        b_match = True
                        break
                if(b_match is False):
                    self._l_project_without_fiche.append(
                        self._list_project_xlsx[idx_p])

    def show_projects_fiches_missing(self):
        for elt in self._l_project_without_fiche:
            print(elt._name_project)

    def show_all_project_data(self):
        for idx, item in enumerate(self._list_project):
            s_projet = "Nom du projet : "
            + self._list_project[idx]._name_project
            s_fiche = " Nom fiches : "
            for idx_l, item_l in enumerate(self._list_project[idx].l_fiches):
                s_fiche = s_fiche + self._list_project[idx].l_fiches[idx_l]
            print(s_projet + s_fiche)

    def _import_existing_suivi(self):
        df = pd.read_csv(self._existing_out_file)
        l_name_project = []
        l_fiches = []
        l_s_t1 = []
        l_s_t2 = []
        l_s_t3 = []
        l_s_t4 = []
        l_n_t1 = []
        l_n_t2 = []
        l_n_t3 = []
        l_n_t4 = []
        l_name_project = df['n_fiche'].values
        l_fiches = df['p_fiche'].values
        l_s_t1 = df['s_t1'].values
        l_s_t2 = df['s_t2'].values
        l_s_t3 = df['s_t3'].values
        l_s_t4 = df['s_t4'].values
        l_n_t1 = df['n_t1'].values
        l_n_t2 = df['n_t2'].values
        l_n_t3 = df['n_t3'].values
        l_n_t4 = df['n_t4'].values

        for i in range(len(l_name_project)):
            n_fiche = l_name_project[i]
            p_fiche = l_fiches[i]
            s_t1 = l_s_t1[i]
            s_t2 = l_s_t2[i]
            s_t3 = l_s_t3[i]
            s_t4 = l_s_t4[i]
            n_t1 = l_n_t1[i]
            n_t2 = l_n_t2[i]
            n_t3 = l_n_t3[i]
            n_t4 = l_n_t4[i]

            s_t = [s_t1, s_t2, s_t3, s_t4]
            n_t = [n_t1, n_t2, n_t3, n_t4]

            t_project = Project(n_fiche, n_t, s_t, None)
            t_project.l_fiches = p_fiche

            self._list_project.append(t_project)

        if(self._Debug):
            print("conteru de la liste année précédente")
            for p in self._list_project:
                print(p._name_project)

        return None

    def import_new_year(self, f_year):
        """ Allow to import a new year to the existing member _list_project.
        return 3 lists continuity, stopped, new prject.

        Parameters
        ----------

        f_year : string
                 filename of the file to import

        Returns
        -------

        l_continu : project list
                    project which contain after processing all project which
                    continu from n-1

        l_close_p : project list
                    project which stopped since n-1

        l_new_proj : project liste
                     new project in year n
        """

        if(self._Debug is True):
            print("Le fichier à importer et merger est : ", f_year)
        df = pd.read_csv(f_year)
        l_name_project = []
        l_n_t1 = []
        l_n_t2 = []
        l_n_t3 = []
        l_n_t4 = []
        l_name_project = df['n_fiche'].values
        l_n_t1 = df['T1'].values
        l_n_t2 = df['T2'].values
        l_n_t3 = df['T3'].values
        l_n_t4 = df['T4'].values

        for i in range(len(l_name_project)):
            n_fiche = l_name_project[i]
            n_t1 = l_n_t1[i]
            n_t2 = l_n_t2[i]
            n_t3 = l_n_t3[i]
            n_t4 = l_n_t4[i]

            s_t = [0, 0, 0, 0]
            n_t = [n_t1, n_t2, n_t3, n_t4]

            t_project = Project(n_fiche, n_t, s_t, None)

            self._l_new_year.append(t_project)

        if(self._Debug):
            print("conteru de la liste de la nouvelle année")
            for p in self._l_new_year:
                print(p._name_project)

        l_continu, l_close_p, l_new_proj = self._merge_2_years()

        return l_continu, l_close_p, l_new_proj

    def _add_empty_field_new_project(self, l_new_proj, nb):
        l_add = []
        for i in range(nb):
            l_add.append(0)

        for proj in l_new_proj:
            proj._nb_heures = l_add + proj._nb_heures
            proj._status = l_add + proj._status

        return l_new_proj

    def process_from_t(self, l_project, last_T):
        """ process a list of project which contain all data after merging
            the full list contain continuity, closed dans new project.
            the present function, place automatically to the last trimester
            status considering last_T.
            Return modified project list

            parameters
            ----------

                l_project : list of project
                            list of project with continuity, closed and new
                            project

                last_T : integer
                         index of last trimester which contain usefull status
            returns
            -------

                l_modif_project : list of project
                                  contain a modified list with update status

        """
        # TODO: faire un prétraitement de la dans la liste de status pour voir
        # s'il y a des NA
        l_modif_project = []
        i_current_status = len(l_project[0]._status) - 1
        for p in l_project:
            if p._status[last_T] != 0:
                if p._status[i_current_status] == -1:
                    p._status[i_current_status] = 10
                else:
                    p._status[i_current_status] = p._status[last_T]
            p_temp = deepcopy(p)
            l_modif_project.append(p_temp)
        return l_modif_project

    def _process_statut(s_nm1):
        s_n = s_nm1
        return s_n

    def _add_empty_field_stopped_project(self, l_stop_proj, nb):
        l_add = []
        for i in range(nb):
            l_add.append(-1)

        for proj in l_stop_proj:
            proj._nb_heures = proj._nb_heures + l_add
            proj._status = proj._status + l_add

        return l_stop_proj

    def _merge_2_years(self):

        l_close_p = []
        l_new_proj = []
        l_continu = []

        T_match = 0.8

        # find close project of n-1
        l_close_p, l_continu = self._extract_from_match(self._list_project,
                                                        self._l_new_year,
                                                        T_match)
        l_new_proj, _ = self._extract_from_match(self._l_new_year,
                                                 self._list_project, T_match)

        for p in self._list_project:
            print(len(p._nb_heures))

        l_close_p = self._add_empty_field_stopped_project(l_close_p,
                                                          len(self._list_project[0]._nb_heures))

        print(len(self._list_project[0]._nb_heures))

        l_new_proj = self._add_empty_field_new_project(l_new_proj,
                                                       len(self._list_project[0]._nb_heures))
        print(len(self._list_project[0]._nb_heures))

        if(self._Debug is True):
            print("Number project without next in 2017", len(l_close_p),)
            for proj in l_close_p:
                print(proj._name_project)
                print(proj._nb_heures)
                print(proj._status)
            print("Number of new project in 2017", len(l_new_proj),)
            for proj in l_new_proj:
                print(proj._name_project)
                print(proj._nb_heures)
                print(proj._status)
            print("Number project which contiue for 2017", len(l_continu))
            for proj in l_continu:
                print(proj._name_project)
                print(proj._nb_heures)
                print(proj._status)

        # get all fiche name
        # self._list_all_fiche(
        #     '/home/cedric/Documents/Conseil/Creative/Data/2017/Fiches/')

        return l_continu, l_close_p, l_new_proj

    def _list_all_fiche(self, dir_path):
        list_files_complet = []
        list_files = []
        for root, dirs, files in os.walk(dir_path):
            for nom_fichier in files:
                list_files_complet.append(os.path.join(root, nom_fichier))
                list_files.append(nom_fichier)

        for idx, iten in enumerate(list_files):
            if(list_files[idx].find("base_mission") != -1):
                del list_files[idx]

        self._liste_fiche = list_files[:]
        self._liste_fiche_full = list_files_complet[:]

        return None

    def _find_fiche_to_proj(self, p_name, T_Ratio):
        idx_maxi = 0
        ratio_maxi = 0
        T = T_Ratio

        for i in range(len(self._liste_fiche)):
            s_file = self._liste_fiche[i]
            ratio = similarity(s_file, p_name)
            if (ratio > ratio_maxi):
                ratio_maxi = ratio
                idx_maxi = i

        if ratio_maxi > T:
            fiche_name = self._liste_fiche[idx_maxi]
        else:
            fiche_name = "Pas match de nom Correct"

        return fiche_name

    def _proc_l_new(self, l):

        l_m = []
        l_s = [0, 0, 0, 0, 0, 0, 0, 0]
        l_h_v = [0, 0, 0, 0]

        for p in l:
            p = Project(l._name_project, l_h_v.extend(p._status), l_s)
            p._AT = True
            l_m.append(p)

        return None

    def _extract_from_match(self, l1, l2, T_match):

        l_no_match = []
        l_match_up = []
        ratio_maxi = 0
        indx_maxi = 0
        l_indx_maxi = []
        l_maxi = []

        print("Pre Proc l1")
        l1p = self._pre_proc_fname(l1)
        print("Pre Proc l2")
        l2p = self._pre_proc_fname(l2)

        for i in range(len(l1p)):
            n_l1 = l1p[i]._name_project
            ratio_maxi = 0
            indx_maxi = 0
            for j in range(len(l2p)):
                n_l2 = l2p[j]._name_project
                ratio = similarity(n_l1, n_l2)
                if(ratio > ratio_maxi):
                    ratio_maxi = ratio
                    indx_maxi = j

            l_indx_maxi.append(indx_maxi)
            l_maxi.append(ratio_maxi)
            if(self._Debug):
                print("MATCH")
                print(n_l1)
                print(l2p[indx_maxi]._name_project)
                print(ratio_maxi)

        # thresold low value match
        if(self._Debug):
            print("Nb indx : ", len(l_indx_maxi), "Nb maxi : ", len(l_maxi))
        for i in range(len(l1)):
            if(l_maxi[i] > T_match):
                # replace with n-1 name
                l1[i]._name_project = l2[l_indx_maxi[i]]._name_project
                # merge n_t
                # TODO: vérifier pourquoi on utilise l2p et non l2
                l_h = l1[i]._nb_heures + l2p[l_indx_maxi[i]]._nb_heures
                l_s = l1[i]._status + [0, 0, 0, 0]
                l_match_up.append(Project(l2[l_indx_maxi[i]]._name_project,
                                          l_h, l_s))
                if(self._Debug):
                    print(l1p[i]._name_project, "Match avec : ",
                          l2p[l_indx_maxi[i]]._name_project)
            else:
                l_temp = deepcopy(l1[i])
                l_no_match.append(l_temp)

        return l_no_match, l_match_up

    def _pre_proc_fname(self, l):
        l_res = []

        for p in l:
            if(p._name_project.find("CA_SPROJ_2016_T1_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2016_T1_", "")
            elif(p._name_project.find("CA_SPROJ_2016_T2_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2016_T2_", "")
            elif(p._name_project.find("CA_SPROJ_2016_T3_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2016_T3_", "")
            elif(p._name_project.find("CA_SPROJ_2016_T4_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2016_T4_", "")
            elif(p._name_project.find("CI_SPROJ_2016_T1_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2016_T1_", "")
            elif(p._name_project.find("CI_SPROJ_2016_T2_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2016_T2_", "")
            elif(p._name_project.find("CI_SPROJ_2016_T3_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2016_T3_", "")
            elif(p._name_project.find("CI_SPROJ_2016_T4_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2016_T4_", "")

            elif(p._name_project.find("2016_T1_Projet_") != -1):
                s_res = p._name_project.replace("2016_T1_Projet_", "")

            elif(p._name_project.find("CA_SPROJ_2017_T1_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2017_T1_", "")
            elif(p._name_project.find("CA_SPROJ_2017_T2_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2017_T2_", "")
            elif(p._name_project.find("CA_SPROJ_2017_T3_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2017_T3_", "")
            elif(p._name_project.find("CA_SPROJ_2017_T4_") != -1):
                s_res = p._name_project.replace("CA_SPROJ_2017_T4_", "")
            elif(p._name_project.find("CI_SPROJ_2017_T1_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2017_T1_", "")
            elif(p._name_project.find("CI_SPROJ_2017_T2_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2017_T2_", "")
            elif(p._name_project.find("CI_SPROJ_2017_T3_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2017_T3_", "")
            elif(p._name_project.find("CI_SPROJ_2017_T4_") != -1):
                s_res = p._name_project.replace("CI_SPROJ_2017_T4_", "")

            p_res = Project(s_res, p._nb_heures, p._status)
            print(p_res._name_project, " ----- ", p._name_project)

            l_res.append(p_res)
        print("Fin Pre Proc")
        return l_res

    def xlxs_to_csv(self, name_xlsx, name_csv):
        i_book = -1
        book = xlrd.open_workbook(name_xlsx)
        print("Ouverture data existant : OK")
        names_sheet = book.sheet_names()

        for idx, item in enumerate(names_sheet):
            if(names_sheet[idx].find("Data") != -1):
                i_book = idx
        if(i_book == -1):
            print("Impossible de retourner les onglet dans le fichie excel")
        all_data = book.sheet_by_index(i_book)
        l_titre = all_data.col_values(0)
        l_T1_s = all_data.col_values(1)
        l_T2_s = all_data.col_values(2)
        l_T3_s = all_data.col_values(3)

        # necessiterat un traitement a l'avenir quand T2, T3, T4 n'existe pas

        for idx, item in enumerate(l_titre):
            l_T_s = [l_T1_s[idx], l_T2_s[idx], l_T3_s[idx]]
            self._list_project_csv.append(Project(l_titre[idx], None, l_T_s))

        # all project xls
        for idx_p, item_p in enumerate(self._list_project):
            ratio_maxi = 0
            # teste si c'est un nouveau project
            if (self._list_project[idx_p]._nb_heures[0] == 0
                    and self._list_project[idx_p]._nb_heures[1] == 0
                    and self._list_project[idx_p]._nb_heures[2] == 0):
                l_vide = [-1, -1, -1]
                self._list_project[idx_p]._status = l_vide
                self._list_project[idx_p]._ratio = -1
                self._list_project[idx_p]._name_csv = "NB heures : "
                + str(self._list_project[idx_p]._nb_heures[0]) + " "
                + str(self._list_project[idx_p]._nb_heures[1]) + " "
                + str(self._list_project[idx_p]._nb_heures[1])
            else:
                for idx_csv, item_csv in enumerate(self._list_project_csv):
                    ratio = similarity(self._list_project[idx_p]._name_project,
                                       self._list_project_csv[idx_csv].
                                       _name_project)
                    if(ratio > ratio_maxi):
                        idx_ratio_maxi = idx_csv
                        ratio_maxi = ratio
                if ratio_maxi != 0:
                    self._list_project[idx_p]._status = self._list_project_csv[
                        idx_ratio_maxi]._status
                    self._list_project[idx_p]._ratio = ratio_maxi
                    self._list_project[idx_p]._name_csv = \
                        self._list_project_csv[idx_ratio_maxi]._name_project

    def proc_trimester(self, last_T):
        T = int(last_T)

        # Clean last T4 semester

        # self._remove_status_T(T+1)

        for idx, item in enumerate(self._list_project):
            # on teste tous les cas de status
            if len(self._list_project[idx]._status) != 0:
                # print str(len(self._list_project[idx]._status))
                # print str(T)
                status = self._list_project[idx]._status[T-1]
                # print status
                # 1 Non eligible exclu de la veille
                if status == 1:
                    self._list_project[idx]._status.append(1)
                # 2 Non eligible suivi 2 fois par an
                elif status == 2:
                    if (T + 1) % 2 == 0:
                        self._liste_a_traiter.append(self._list_project[idx])
                    else:
                        self._list_project[idx]._status.append(2)
                # 3 Non determine element favable suivi stric
                elif status == 3:
                    if(self._list_project[idx]._nb_heures[T] != 0):
                        self._liste_a_traiter.append(self._list_project[idx])
                    else:
                        self._list_project[idx]._status.append(3)
                # 4 Eligible et va etre ecrit
                # 5 Rescrit obtenu ou en cours
                elif status == 5:
                    if(self._list_project[idx]._nb_heures[T] != 0):
                        self._liste_a_traiter.append(self._list_project[idx])
                    else:
                        self._list_project[idx]._status.append(5)
                # 6 Continuite
                elif status == 6:
                        self._list_project[idx]._status.append(6)

                # 7 Petite Fiche existante
                elif status == 7:
                        self._list_project[idx]._status.append(7)
                elif status == 0 or status == -1:
                    # combien append
                    val_ap = 4 - len(self._list_project[idx]._status)
                    for i in range(val_ap):
                        self._list_project[idx]._status.append(0)
                    if(self._list_project[idx]._nb_heures[T] != 0):
                        self._liste_a_traiter.append(self._list_project[idx])
            else:
                # add 0 for unexisting Tri
                for i in range(T):
                    self._list_project[idx]._status.append(-1)
                # pas de statut soit nouveau donc a traiter
                self._list_nouveau.append(self._list_project[idx])
                self._liste_a_traiter.append(self._list_project[idx])

    def show_proj_a_traiter(self):
        print("Projet a traiter")
        for idx, item in enumerate(self._liste_a_traiter):
            print("Nom de projet : " +
                  self._liste_a_traiter[idx]._name_project)
            if len(self._liste_a_traiter[idx]._l_fiches) != 0:
                    print("open fiche : " +
                          self._liste_a_traiter[idx]._l_fiches[0])
                    self._open_pdf(self._liste_a_traiter[idx]._l_fiches[0])

    def _remove_status_T(self, T):
        for idx, item in enumerate(self._list_project):
            if len(self._list_project[idx]._status) != 0:
                print(str(len(self._list_project[idx]._status)))
                del self._list_project[idx]._status[T-1]

    def _open_pdf(self, name):
        subprocess.call(["okular", name])

    def proc_proj_a_traiter(self):
        print("Projet a traiter")
        for idx, item in enumerate(self._liste_a_traiter):
            print("Nom de projet : " +
                  self._liste_a_traiter[idx]._name_project)
            if len(self._liste_a_traiter[idx]._l_fiches) != 0:
                    print("open fiche : " +
                          self._liste_a_traiter[idx]._l_fiches[0])
                    self._open_pdf(self._liste_a_traiter[idx]._l_fiches[0])
                    print("Dernier status : "
                          + str(self._liste_a_traiter[idx]._status[2]))
                    try:
                        mode = int(input('Enter new status'))
                        print("Valeur saisie : " + str(mode))
                    except:
                        print("Not a number")
                    # on cherche le projet dans tt les fiches
                    for idx_l, item_l in enumerate(self._list_project):
                        # print str(idx_l)
                        ratio = similarity(
                            self._liste_a_traiter[idx]._name_project,
                            self._list_project[idx_l]._name_project)
                        # print ratio
                        if ratio == 1:
                            # virer la constante ici
                            if len(self._list_project[idx_l]._status) < 4:
                                self._list_project[idx_l]._status.append(mode)
                            else:
                                self._list_project[idx_l]._status[3] = mode
                            print(self._list_project[idx_l]._status[3])
                            break

    def export_l_csv(self, name, l):
        with open(name, 'w') as fl:
            print("nombre projet export : " + str(len(l)))
            for idx, item in enumerate(l):
                a_ecrire = l[idx]._name_project + ","
                if len(l[idx]._l_fiches) != 0:
                    a_ecrire = a_ecrire + l[idx]._l_fiches[0] + ","
                else:
                    a_ecrire = a_ecrire + " " + ","
                if len(l[idx]._status) != 0:
                    for idx_s, item_s in enumerate(l[idx]._status):
                        a_ecrire = a_ecrire + str(l[idx]._status[idx_s]) + ","
                else:
                    a_ecrire = a_ecrire + " " + "," + " " + "," + " " + ","
                    + " " + ","
                # fl.write("%s\n" % a_ecrire)
                if len(l[idx]._nb_heures) != 0:
                    for idx_s, item_s in enumerate(l[idx]._nb_heures):
                        a_ecrire = a_ecrire + str(l[idx]._nb_heures[idx_s]) \
                            + ","
                else:
                    a_ecrire = a_ecrire + " " + "," + " " + "," + " " + ","
                    + " " + ","
                if self._Debug is True:
                    print('Type of string to write : ', type(a_ecrire))
                fl.write(a_ecrire + '\n')
                a_ecrire = ''
            fl.close()

    def export_l_html(self, name, l):
        with open(name, 'wb') as fl:
            print("nombre projet export : " + str(len(l)))
            fl.write("<!DOCTYPE html>")
            fl.write("<html>")
            fl.write("<head>")
            fl.write("<style> \n table, th, td {border: 1px solid black;}")
            fl.write("</style>")
            fl.write("</head>")
            fl.write("<body>")
            fl.write("<table>")
            fl.write("<tr>")
            for idx, item in enumerate(l):
                if l[idx]._name_project != "":
                    a_ecrire = "<th>" + l[idx]._name_project + "</th>"
                else:
                    a_ecrire = "<th>Pas de nom de projet</th>"
                if len(l[idx]._l_fiches) != 0:
                    a_ecrire = a_ecrire + "<th>" + "<a href = '"
                    + l[idx]._l_fiches[0].replace(
                        "/home/cedric/Documents/Conseil/Creative", "..")
                    + "'>" + l[idx]._l_fiches[0].replace(
                        "/home/cedric/Documents/Conseil/Creative", "..")
                    + "</a>" + "</th>"
                else:
                    a_ecrire + a_ecrire + "<th>Pas de fichier dispo</th>"
                if (len(l[idx]._status) != 0 and len(l[idx]._nb_heures) != 0):
                    cumul = 0
                    for idx_s, item_s in enumerate(l[idx]._status):
                        status = l[idx]._status[idx_s]
                        s_bg = self._status_to_color(status)
                        if(idx_s < len(l[idx]._nb_heures)):
                            a_ecrire = a_ecrire + "<td bgcolor=" + s_bg + ">"
                            + str(l[idx]._nb_heures[idx_s]) + "</td>"
                            cumul = cumul + l[idx]._nb_heures[idx_s]
                        else:
                            a_ecrire = a_ecrire + "<td bgcolor=" + s_bg + ">"
                            + str(-1) + "</td>"
                    a_ecrire += "<th>" + str(cumul) + "</th>"
                    cumul = 0

                else:
                    a_ecrire = a_ecrire + "<th>0</th>"
                    + "<th>0</th>" + "<th>0</th>" + "<th>0</th>"
                # print a_ecrire
                # print str(idx)
                fl.write("%s\n" % a_ecrire)
                fl.write("</tr>")
                a_ecrire = ''
            fl.write("</tr>")
            fl.write("</table")
            fl.write("</html>")
            fl.close()

    def _status_to_color(self, status):
        if(status == 1):
            return "#A0A0A0"
        elif(status == 2):
            return "#FF0000"
        elif(status == 3):
            return "#FF8000"
        elif(status == 4):
            return "#FFFF00"
        elif(status == 5):
            return "#00FF00"
        elif(status == 6):
            return "#0000FF"
        elif(status == 7):
            return "#7f00FF"
        elif (status == -1):
            return "#007FA0"
        else:
            return "#FFFFFF"

    def clean_l_proj(self, l):
        for idx, item in enumerate(l):
            if l[idx]._name_project == "":  # or l[idx]._l_fiches[0] == "" :
                del l[idx]
