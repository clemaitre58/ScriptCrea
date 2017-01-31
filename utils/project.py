class Project :
    def __init__(self, name_project, l_heures = None) :
        self._name_project = name_project
        self._nb_heures = []
        self.status = []
        self.add_fiche = []

        if (l_heures != None):
            self._nb_heures = l_heures

#    def add_trimester() :



