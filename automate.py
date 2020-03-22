def newDict(a):

    # Populating the initial dictionary with empty list
    newDict = {}
    for i in a:
        newDict[i] = []
    return newDict


def padDict(d, pad=""):
    length = len(d['Serial No'])
    for key, value in d.items():
        if len(value) < length:
            d[key].append(pad)
    return d


def convert(inputFilePath='inputdata.xlsx', outputFilePath='outputData.xlsx', a=[
    'Serial No',
    'Name (as per NRIC)',
    'NRIC',
    'Address (as per NRIC)',
    'Mailing Address',
    'Contact Number (Residence)',
    'Contact Number (Mobile)',
    'Email address',
    'SSIC',
    'UEN number',
    'Mode of payment',
    'Name (as per bank book)',
    'Bank name',
        'Bank account number']):

    import pandas as pd

    rowIndex = 0
    outDict = newDict(a)

    # Get Excel Book
    xls = pd.ExcelFile(inputFilePath)
    datasheetNames = xls.sheet_names  # To get the list of the name of datasheet

    # Loop every datasheet in the book
    for datasheetName in datasheetNames:
        # setup dictionary for the output file
        # print(outDict)
        outDict['Serial No'].append(0)
        rowIndex += 1

        df = xls.parse(datasheetName, header=None)
        convDict = df.to_dict()
        for index, key in convDict[0].items():
            if outDict.get(str(convDict[0][index]).strip(), [-1]) != [-1]:
                outDict[str(convDict[0][index]).strip()
                        ].append(convDict[1][index])

        outDict = padDict(outDict)

    outDf = pd.DataFrame(data=outDict, index=None)
    with pd.ExcelWriter('output.xlsx', mode='w') as writer:
        outDf.to_excel(writer, sheet_name='FINAL')
    return outDf


if __name__ == "__main__":
    import sys
    args = sys.argv

    outputFilePath = None
    inputFilePath = None

    helpMessage = """
        Arguments: 
            -o or -O    output file path
            -i or -I    input file path

        Note:
            The output file will be overwritten.\

        Example:
            $ python ./automate.py -i inputdata.xlsx -o outputdata.xlsx
        """

    if len(args) == 1:
        print(helpMessage)
        sys.exit(0)

    if "-o" in args or "-O" in args:
        try:
            outputFilePath = args[args.index("-o") + 1]
        except:
            outputFilePath = args[args.index("-O") + 1]

    if "-i" in args or "-I" in args:
        try:
            inputFilePath = args[args.index("-i") + 1]
        except:
            inputFilePath = args[args.index("-I") + 1]

    if not (outputFilePath and inputFilePath):
        print(helpMessage)

    print("processing {}. Output file can be found at {}".format(inputFilePath, outputFilePath))
    convert(inputFilePath, outputFilePath)
    print("Done")

