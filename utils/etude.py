import csv
import os
import xlrd
import subprocess
from utils.project import *
from utils.str import *
class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file = None):
        self._list_project = []
        self._list_project_csv = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        self._data_path = self._dir_proj + '/Data'
        self._l_project_without_fiche = []
        self._list_nouveau = []
        self._list_conti = []
        self._list_petite_fiche = []
        self._liste_a_traiter = []

        if (self._existing_out_file == None):
            self._existing_out_file = self._dir_proj + '/Suivi/suivi.csv'
            self._create_empty_ofile(self._existing_out_file)
        else :
            self._import_existing_suivi(self._existing_out_file)


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
            self._parse_xls_base_mission(name_dir_to_list)
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

    def _parse_xls_base_mission(self, dir_path):
        i_book_atl = -1
        i_book_ing = -1
        name_base_mission = dir_path + "base_mission.xlsx"
        book = xlrd.open_workbook(name_base_mission)
        print "Base mission : OK" #+ name_base_mission
        names_sheet = book.sheet_names()
        
        for idx, item in enumerate(names_sheet):
            if(names_sheet[idx].find("Suivi par Fiches Ingenierie") !=-1):
                i_book_ing = idx
            if(names_sheet[idx].find("Suivi par Fiches Atlantique") !=-1):
                i_book_atl = idx
        if(i_book_atl == -1 and i_book_ing ==-1):
            print "Impossible de retourner les onglet dans le fichie excel"

        print "Position onglet atlantique : " + str(i_book_atl) + "\n" + "Position onglet Ingenierie : " + str(i_book_ing)
        l_idx = [i_book_atl, i_book_ing]
        for idx_b, item_b in enumerate(l_idx):

            all_data = book.sheet_by_index(item_b)
            l_titre = all_data.col_values(0)
            l_T1 = all_data.col_values(1)
            l_T2 = all_data.col_values(2)
            l_T3 = all_data.col_values(3)
            l_T4 = all_data.col_values(4)
    
            # necessiterat un traitement a l'avenir quand T2, T3, T4 n'existe pas
            
            for idx, item in enumerate(l_titre):
                l_T = [l_T1[idx] , l_T2[idx], l_T3[idx], l_T4[idx]]
                self._list_project.append(Project(l_titre[idx], l_T, None))

    def match_fiche_data(self):
        b_match = False
        if(self._list_project != None 
                and self._all_tri_fiches != None):
            for idx_p, item_p in enumerate(self._list_project):
                b_match = False
                for idx_f, item_f in enumerate(self._all_tri_fiches):
                    if(self._all_tri_fiches[idx_f].find(
                        self._list_project[idx_p]._name_project) != -1):
                        self._list_project[idx_p].add_fiche(self._all_tri_fiches[idx_f])
                        b_match = True
                        break
                if(b_match == False):
                    self._l_project_without_fiche.append(self._list_project[idx_p])

    def show_projects_fiches_missing(self):
        for elt in self._l_project_without_fiche:
            print elt._name_project
    
    def show_all_project_data(self):
        for idx, item in enumerate(self._list_project):
            s_projet = "Nom du projet : " + self._list_project[idx]._name_project
            s_fiche = " Nom fiches : "
            for idx_l, item_l in enumerate(self._list_project[idx].l_fiches):
                s_fiche = s_fiche + self._list_project[idx].l_fiches[idx_l]
            print s_projet + s_fiche

    def _import_existing_suivi(self):
        return None

    def xlxs_to_csv(self, name_xlsx, name_csv):
        i_book = -1
        book = xlrd.open_workbook(name_xlsx)
        print "Ouverture data existant : OK" 
        names_sheet = book.sheet_names()
        
        for idx, item in enumerate(names_sheet):
            if(names_sheet[idx].find("Data") !=-1):
                i_book = idx
        if(i_book==-1):
            print "Impossible de retourner les onglet dans le fichie excel"
        all_data = book.sheet_by_index(i_book)
        l_titre = all_data.col_values(0)
        l_T1_s = all_data.col_values(1)
        l_T2_s = all_data.col_values(2)
        l_T3_s = all_data.col_values(3)
    
            # necessiterat un traitement a l'avenir quand T2, T3, T4 n'existe pas
            
        for idx, item in enumerate(l_titre):
            l_T_s = [l_T1_s[idx] , l_T2_s[idx], l_T3_s[idx]]
            self._list_project_csv.append(Project(l_titre[idx], None, l_T_s))
        
        # all project xls
        for idx_p, item_p in enumerate(self._list_project):
            ratio_maxi = 0
            for idx_csv, item_csv in enumerate(self._list_project_csv):
                print "Nom projet : " + self._list_project[idx_p]._name_project + " Nom CSV : " + self._list_project_csv[idx_csv]._name_project
                ratio = similarity(self._list_project[idx_p]._name_project, 
                            self._list_project_csv[idx_csv]._name_project)
                print "ratio : " + str(ratio)
                if (ratio > ratio_maxi):
                    idx_ratio_maxi = idx_csv
                    ratio_maxi = ratio
            if ratio_maxi != 0.0:
                self._list_project[idx_p]._status = self._list_project_csv[idx_ratio_maxi]._status

    def proc_trimester(self, last_T):
        T = int(last_T)

        # Clean last T4 semester

        #self._remove_status_T(T+1)

        for idx, item in enumerate(self._list_project):
            # on teste tous les cas de status
            if len(self._list_project[idx]._status) != 0 :
                print str(len(self._list_project[idx]._status))
                print str(T)
                status = self._list_project[idx]._status[T-1]
                print status
                # 1 Non eligible exclu de la veille
                if status == 1 :
                    self._list_project[idx]._status.append('1')
                # 2 Non eligible suivi 2 fois par an
                elif status == 2 :
                    if (T + 1) % 2 ==  0 :
                        self._liste_a_traiter.append(self._list_project[idx])
                    else:
                        self._list_project[idx]._status.append('2')
                # 3 Non determine element favable suivi stric
                elif status == 3:
                    self._liste_a_traiter.append(self._list_project[idx])
                # 4 Eligible et va etre ecrit
                # 5 Rescrit obtenu ou en cours
                elif status == 5 :
                    self._liste_a_traiter.append(self._list_project[idx])
                # 6 Continuite
                elif status == 6:
                    if (T + 1) == 4:
                        self._liste_a_traiter.append(self._list_project[idx])
                        self._list_conti.append(self._list_project[idx])
                    else:
                        self._list_project[idx]._status.append('6')

                # 7 Petite Fiche existante
                if status == 7 :
                    if (T+ 1) == 4 :
                        self._liste_a_traiter.append(self._list_project[idx])
                        self._list_petite_fiche.append(self._list_project[idx])
                    else :
                        self._list_project[idx]._status.append('7')
            else:
                # add 0 for unexisting Tri
                for i in range (T):
                    self._list_project[idx]._status.append('0')
                # pas de statut soit nouveau donc a traiter
                self._list_nouveau.append(self._list_project[idx])
                self._liste_a_traiter.append(self._list_project[idx])

    def show_proj_a_traiter(self):
        print "Projet a traiter"
        for idx, item in enumerate(self._liste_a_traiter):
            print "Nom de projet : " + self._liste_a_traiter[idx]._name_project
            if len(self._liste_a_traiter[idx]._l_fiches) != 0:
                    print "open fiche : " + self._liste_a_traiter[idx]._l_fiches[0]
                    self._open_pdf(self._liste_a_traiter[idx]._l_fiches[0])

    def _remove_status_T(self, T):
        for idx, item in enumerate(self._list_project):
            if len(self._list_project[idx]._status) !=0:
                print str(len(self._list_project[idx]._status))
                del self._list_project[idx]._status[T-1]

    def _open_pdf(self, name):
        subprocess.call(["okular", name])

    def proc_proj_a_traiter(self):
        print "Projet a traiter"
        for idx, item in enumerate(self._liste_a_traiter):
            print "Nom de projet : " + self._liste_a_traiter[idx]._name_project
            if len(self._liste_a_traiter[idx]._l_fiches) != 0:
                    print "open fiche : " + self._liste_a_traiter[idx]._l_fiches[0]
                    self._open_pdf(self._liste_a_traiter[idx]._l_fiches[0])
                    try :
                        mode = int(raw_input('Enter new status'))
                        print "Valeur saisie : " + str(mode)
                    except :
                        print "Not a number"


    def export_l_csv(self, name ,l):
        with open(name, 'wb') as fl:
            print "nombre projet export : " + str(len(l))
            for idx, item in enumerate(l):
                a_ecrire = l[idx]._name_project + ","
                if len (l[idx]._l_fiches) != 0:
                    a_ecrire = a_ecrire + l[idx]._l_fiches[0] + ","
                else:
                    a_ecrire + a_ecrire + " "
                if len(l[idx]._status) != 0:
                    for idx_s, item_s in enumerate(l[idx]._status):
                        a_ecrire = a_ecrire + str(l[idx]._status[idx_s]) + ","
                else:
                    a_ecrire = a_ecrire + " " + "," + " " + "," + " " + "," + " " + ","
                #fl.write("%s\n" % a_ecrire)
                if len(l[idx]._nb_heures) != 0:
                    for idx_s, item_s in enumerate(l[idx]._nb_heures):
                        a_ecrire = a_ecrire + str(l[idx]._nb_heures[idx_s]) + ":"
                else:
                    a_ecrire = a_ecrire + " " + "," + " " + "," + " " + "," + " " + ","
                fl.write("%s\n" % a_ecrire)
                a_ecrire = ''
            fl.close()

    def export_l_html(self, name, l):
        with open(name, 'wb') as fl:
            print "nombre projet export : " + str(len(l))
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
                a_ecrire = "<th>" + l[idx]._name_project + "</th>"
                if len (l[idx]._l_fiches) != 0:
                    a_ecrire = a_ecrire + "<th>" + l[idx]._l_fiches[0] + "</th>"
                else:
                    a_ecrire + a_ecrire + "<th> </th>"
                if (len(l[idx]._status) != 0) and len(l[idx]._nb_heures != 0):
                    for idx_s, item_s in enumerate(l[idx]._status):
                        status = l[idx]._status[idx_s]
                        s_bg = _status_to_color(status)
                        a_ecrire = a_ecrire + "<td bgcolor=" + s_bg + ">" + str(l[idx]._nb_heures[idx_s]) + "</td>"
                else:
                    a_ecrire = a_ecrire + "<th> </th>"+ "<th> </th>"+ "<th> </th>"+ "<th> </th>"                 fl.write("%s\n" % a_ecrire)
                a_ecrire = ''
            fl.write("</tr>")
            fl.write("</table")
            fl.write("</html>")
            fl.close()
        
    def _status_to_color(status):
        if(status == 1):
            return "#000000"
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
        else:
            return "#FFFFFF"
