import os

modeladdr = 'C:\\Users\\SJ\\PycharmProjects\\finalfile\\__pycache__'


def filename(modeladdr_):
    modnamelist = os.listdir(modeladdr_)
    modnamelisttem = ['' for _ in range(len(modnamelist))]
    for i in range(len(modnamelist)):
        if modnamelist[i].endswith('.cpython-39.opt-2.pyc'):
            modnamelisttem[i] = modnamelist[i]
            os.rename(modeladdr + r'\%s' % modnamelisttem[i],
                      modeladdr + r'\%s' % modnamelist[i].replace('.cpython-39.opt-2.pyc', '.pyc'))


def main():
    filename(modeladdr)


if __name__ == '__main__':
    main()
