import re

trash = ['tabs', 'inj', 'bottle', 'syp', 'bot', 'bott', 'cap', 'doses', 'with', 'ml', 'mg', 'in', 'methyl', \
         'containing', 'antibiotic', 'sodium', 'chloride', 'fluoride', 'phosphate', 'without', \
         'chloride', 'ammonium', 'citrate', 'adrenaline', 'gluconate', 'propionate', 'absorbent', \
         'unmedicated', 'sulphate', 'eye drops', 'lactate', 'disposable', 'lignocaine']

contraVals = [['tab', 'tabs', 'inj', 'bottle', 'syp', 'bot', 'bott', 'cap', 'drops', 'needles', 'ointment'],
              ['sodium', 'chloride', 'fluoride', 'phosphate']]


def isSimilar(v1, v2):
    if v1.lower() == v2.lower():
        return 0
    if v1.replace(' ', '') == v2.replace(' ', ''):
        return 0
    else:
        a = re.findall(r"[\w]+", v1.lower())
        b = re.findall(r"[\w]+", v2.lower())
        a = list(set(a) - set(trash))
        b = list(set(b) - set(trash))
        if a == b:
            return 0
        else:
            for _ in a[1:len(a) - 2]:
                if _.isalpha():
                    for p in b:
                        if _ == p and len(_) > 5:
                            # print(' '.join(a),' '.join(b))
                            return 1
                            # TODO: Make the function to ask matching seperately
    return 2
