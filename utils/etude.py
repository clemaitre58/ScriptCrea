import csv
import os
import xlrd
from utils.project import *

class EtudeFiches:
    def __init__(self, dir_proj, existing_out_file = None):
        self._list_project = []
        self._dir_proj = dir_proj
        self._existing_out_file = existing_out_file
        self._data_path = self._dir_proj + '/Data'
        self._l_project_without_fiche = []
        
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
                self._list_project.append(Project(l_titre[idx], l_T))

    def match_fiche_data(self):
        b_match = False
        if(self._list_project != None 
                and self._all_tri_fiches != None):
            for idx_p, item_p in enumerate(self._list_project):
                b_match = False
                #print str(idx_p)
                #print self._list_project[idx_p]._name_project
                for idx_f, item_f in enumerate(self._all_tri_fiches):
                    #print self._all_tri_fiches[idx_f]
                    #if(self._list_project[idx_p]._name_project.find(
                    #    self._all_tri_fiches[idx_f]) != -1):
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
