class Project :
    def __init__(self, name_project, l_heures = None) :
        self._name_project = name_project
        self._nb_heures = []
        self.status = []
        self._l_fiches = []

        if (l_heures != None):
            self._nb_heures = l_heures

#    def add_trimester() :

    def add_fiche(self, name_fiche):
        self._l_fiches.append(name_fiche)

    @property
    def l_fiches(self):
        return self._l_fiches

    @l_fiches.setter
    def l_fiches(self, value):
        self._l_fiches = value

