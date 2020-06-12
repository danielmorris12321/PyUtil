class direcError(Exception):
    """ Exception raised for errors in creating a directory

    Attributes: 
        direc - Path of directory where error has occured

    """

    def __init__(self, direc):
        self.direc = direc
        self.message = "Error in creating the following directory: " 
    def __str__(self):
        return (self.message + self.direc)

class multipleFilesError(Exception):
    """ Exception raised for errors in creating a directory

    Attributes: 
        direc - Path of directory where error has occured

    """

    def __init__(self, direc, pattern):
        self.direc = direc
        self.pattern = pattern
        self.message = "Multiple files matching: " + str(pattern) + " in " + direc 
    def __str__(self):
        return (self.message)