class TriHName(object):
    """ Object which allow to store trimister (status cols) name and and hour
    name in data csv file
     """
    def __init__(self, tri_name, h_name):
        """ Create the oblect with agrs tri_name and h_name
        which are string"""
        self._tri_name = tri_name
        self._h_name = h_name
