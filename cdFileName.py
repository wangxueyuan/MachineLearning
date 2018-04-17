class TEST():
    def __init__(self, path):
        assert isinstance(path, object)
        self.path = path

    def __readfile__( path):
        """

        :rtype : object
        """
        listnames = os.listdir(path)
        for name in listnames:
            if(name!=name.zfill(3)):
                oldname=path+"\\"+name
                newname=path+"\\"+name.zfill(3)
                os.rename(oldname, newname)



if __name__ == '__main__':
    path = 'C:\\Users\\Administrator\\Desktop\\test'
    TEST.__readfile__(path)

