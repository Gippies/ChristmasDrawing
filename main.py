import csv


def match_people():
    with open("Testfile.csv", newline="") as csv_file:
        reader = csv.reader(csv_file)
        people = {}

        # First, extract the people from the csv and build a dictionary with people they're not to be paired with.
        for line in reader:
            if line[0] in people.keys():
                print("ERROR, ERROR. There are duplicate people listed in the provided file. Exiting...")
                return
            people[line[0]] = []
            for person in line[1:]:
                people[line[0]].append(person)
        print(people)


if __name__ == '__main__':
    match_people()
