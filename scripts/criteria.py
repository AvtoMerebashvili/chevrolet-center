# this is criteria setter file


def setCriteria(criteria):
    with open("../data/criteria.txt", "w") as file:
        file.write(str(criteria))