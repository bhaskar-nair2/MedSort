import re


trash = ['tabs', 'inj', 'bottle', 'syp', 'bot', 'bott', 'cap', 'doses','with','ml','mg','in', 'methyl',\
         'containing','antibiotic', 'sodium', 'chloride', 'fluoride', 'phosphate','without', \
         'chloride', 'ammonium', 'citrate','adrenaline','gluconate','propionate','absorbent',\
         'unmedicated','sulphate','eye drops','lactate','disposable','lignocaine']

contraVals=[['tab', 'tabs' ,'inj', 'bottle', 'syp', 'bot', 'bott', 'cap','drops','needles','ointment'],['sodium', 'chloride', 'fluoride', 'phosphate']]

def contradict(a,b):
    for _ in a:
        for i in contraVals:
            if _ in i:
                for h in b:
                    if h in i and h!=_:
                        return True #Yes it does Contradict
                    if h==_:
                        return False #No Contradiction
                    return True #Contradicts
    for _ in b:
        for i in contraVals:
            if _ in i:
                for h in a:
                    if h in i and h!=_:
                        return True
                    if h==_:
                        return False
                    return True
    return False

def remove(val):
    return list(set(val) - set(trash))

def isSimilar(v1, v2):
    if v1.lower() == v2.lower():
        return True
    else:
        a = re.findall(r"[\w]+", v1.lower())
        b = re.findall(r"[\w]+", v2.lower())
        if contradict(a,b):
            return False # Stop cause it contradicts
        a = remove(a)
        b = remove(b)
        if a == b:
            return True
        else:
            for _ in a[1:len(a) - 2]:
                if _.isalpha():
                    for p in b:
                        if _ == p and len(_)>5:
                            d = input("Is\n " + v1 + "\nsimilar to\n" + v2+"\n Similar: "+_+" : "+p + "\n [y/n]: ")
                            if d == 'y':
                                return True

    return False




