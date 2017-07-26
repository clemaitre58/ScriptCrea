class Project:
    def __init__(self, name_project, l_heures=None, l_status=None):
        self._name_project = name_project
        self._nb_heures = []
        self._status = []
        self._l_fiches = []
        self._l_fiches_rel = []
        self._name_csv = ""
        self._ratio = 0

        if (l_heures is not None):
            self._nb_heures = l_heures
        if (l_status is not None):
            self._status = l_status

#    def add_trimester() :

    def add_fiche(self, name_fiche):
        self._l_fiches.append(name_fiche)

    @property
    def l_fiches(self):
        return self._l_fiches

    @l_fiches.setter
    def l_fiches(self, value):
        self._l_fiches = value
