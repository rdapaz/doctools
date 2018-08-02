
startCol = 'G'
rowNumber = 126
sheetName = 'Forecast'

def columnLetterGivenNumber(iVal):
    retVal = None
    if iVal <= 26:
        retVal = chr(64+iVal)
    else:
        m = int(iVal/26)
        n = iVal - m*26
        retVal = f'{columnLetterGivenNumber(m)}{columnLetterGivenNumber(n)}' 
    return retVal

def columnNumberGivenLetter(colLetter):
    for i in range(1, 1000):
        v = columnLetterGivenNumber(i)
        if v == colLetter:
            return i

start = columnNumberGivenLetter(startCol)
print(start)
arr = []

for i in range(start, start+24):
    columnLetter = columnLetterGivenNumber(i)
    arr.append('=\'{}\'!{}{}'.format(sheetName, columnLetter, rowNumber))

for entry in arr:
    print(entry)

