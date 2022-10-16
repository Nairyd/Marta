from Util.Helper import getValuesFromExcel, checkKasualie


if __name__ == "__main__":

    myDict = getValuesFromExcel()

    print(myDict)
    x = myDict.get("täufling")
    print(x.lastname)
    checkKasualie(myDict)