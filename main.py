def match_people():
    f = open("Testfile.txt", "r")
    for line in f:
        print(line)
    f.close()


if __name__ == '__main__':
    match_people()
