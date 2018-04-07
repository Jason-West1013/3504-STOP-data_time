import csv
import math
import numpy as np
import matplotlib.pyplot as plt

femaleQuestionDistributions = {}
maleQuestionDistributions = {}
hellingerDistances = {}
genderZScores = {}

# Computes the distributions of male and female responses for the 26 questions
# changed for .csv files
def extractDistributions(column):
    csvFile = open("mini-project2-data-v3.csv", "r")
    reader = csv.reader(csvFile)
    femaleSampleSize = 0
    maleSampleSize = 0

    maleQuestionResponses = []
    femaleQuestionResponses = []
    maleResponses = []
    femaleResponses = []

    for row in reader:
        if row[column] != '0':
            if row[26] == '1':
                femaleResponses.append(int(row[column]))
                femaleSampleSize += 1
            elif row[26] == '2':
                maleResponses.append(int(row[column]))
                maleSampleSize += 1
    
    csvFile.close()

    for i in range(5):
        femaleQuestionResponses.append(femaleResponses.count(i + 1))
        maleQuestionResponses.append(maleResponses.count(i + 1))
    
    femaleDistributions = np.divide(femaleQuestionResponses, len(femaleResponses))
    maleDistributions = np.divide(maleQuestionResponses, len(maleResponses))

    return femaleDistributions, maleDistributions, femaleSampleSize, maleSampleSize

# Calculates the Mean and Variance and then the Z-Score for the male and female distributions
def calculateZTest(responseNum, femaleDist, maleDist, femaleSampleSize, maleSampleSize):
    femaleEx = sum(np.multiply(responseNum, femaleDist))
    maleEx = sum(np.multiply(responseNum, maleDist))

    femaleDoubleEx = sum(np.multiply(np.square(responseNum), femaleDist))
    maleDoubleEx = sum(np.multiply(np.square(responseNum), maleDist))

    femaleVar = femaleDoubleEx - femaleEx**2
    maleVar = maleDoubleEx - maleEx**2

    dividend = maleEx - femaleEx
    divisor = math.sqrt((maleVar / maleSampleSize) + (femaleVar / femaleSampleSize))
    return dividend / divisor

# Calculates the Hellinger Distances between the male and female distributions 
def calculateHellbringerDistances(maleDist, femaleDist):
    divSqRtTwo = 1 / math.sqrt(2)

    tmp = math.sqrt(np.sum(np.square(np.subtract(np.sqrt(maleDist), np.sqrt(femaleDist)))))
    hellin = tmp * divSqRtTwo
    return hellin

# Draws a bar graph for a question showing the male and female distributions 
def drawBarGraph(maleDist, femaleDist, questionTitle):
    nGroups = 5
    fig, p = plt.subplots()
    index = np.arange(nGroups)
    barWidth = 0.35

    maleBar = p.bar(index, tuple(maleDist), barWidth, color='g', label='Men')
    femaleBar = p.bar(index + barWidth, tuple(femaleDist), barWidth, color='b', label='Female')

    p.set_xlabel('Responses')
    p.set_ylabel('Distributions')
    p.set_title("{0} Distributions for Men and Women".format(questionTitle))
    p.set_xticks(index + barWidth / 2)
    p.set_xticklabels(('1', '2', '3', '4', '5'))
    p.legend()

    fig.tight_layout()
    plt.show()

# Writes data to file, intended to be used for the Hellinger Distances
def writeDataToFile(filename, data):
    f = open(filename, "w")
    for key, value in data.items():
        if not isinstance(value, list):
            f.write("%s: %s\n" % (key, round(value, 4)))
        else:
            f.write("%s\n" % key.upper())
            for i in value:
                f.write("%s\n" % i)
            f.write("\n")

# Main
response = list(range(1, 6))
significanceLevel05 = {'accept':[], 'reject':[]}
significanceLevel01 = {'accept':[], 'reject':[]}


# Main loop for the 26 questions
for i in range(26):
    colLetter = chr(65 + i)            
    questionName = "Q%d" % (i + 1)

    # computation
    femaleProb, maleProb, femaleSampleSize, maleSampleSize = extractDistributions(i)
    hellinDist = calculateHellbringerDistances(maleProb, femaleProb)
    zScore = calculateZTest(response, femaleProb, maleProb, femaleSampleSize, maleSampleSize)
    drawBarGraph(femaleProb, maleProb, questionName)

    # Data for gender Hellinger distances and Z-Scores
    hellingerDistances.update({questionName:hellinDist})
    genderZScores.update({questionName:zScore})

minHellinDistKey = min(hellingerDistances, key=hellingerDistances.get)
maxHellinDistKey = max(hellingerDistances, key=hellingerDistances.get)

for key, value in genderZScores.items():
    score = float(value)
    if score < -1.96 or score > 1.96:
        significanceLevel05.get('reject').append(key)
    else:
        significanceLevel05.get('accept').append(key)
    if score < -2.576 or score > 2.576:
        significanceLevel01.get('reject').append(key)
    else:
        significanceLevel01.get('accept').append(key)

writeDataToFile("gender_hellinger.txt", hellingerDistances)
writeDataToFile("01_z_scores.txt", significanceLevel01)
writeDataToFile("05_z_scores.txt", significanceLevel05)
print("The question with the maximum Hellinger Distance is %s, and the question with the minimum is %s." % (maxHellinDistKey, minHellinDistKey))
