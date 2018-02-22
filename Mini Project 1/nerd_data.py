import math
from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt

# opening .xlsx file
wb = load_workbook('mini-project1-data.xlsx')                                                                            
ws = wb[wb.sheetnames[0]]                                          
ws = wb.active                                              

# Global Variables
questionDistributions = {}
questionMeanVarEnt = {}
questionTitles = []

''' Parses each column using openpyxl, then computes the
    and returns the response distribution of the specific 
    question. '''
def distributionOfColumn(columnLabel):
    questionResponses = []
    responses = []

    # parse .xlsx file and add to list
    for row in range(2, ws.max_row + 1):
        for column in columnLabel:
            cellName = "{}{}".format(column, row)
            if ws[cellName].value != 0:                             # omit the 0s from the data set
                responses.append(ws[cellName].value)  

    # populate list with rates per response
    for i in range(5):
        questionResponses.append(responses.count(i + 1))
    
    return np.divide(questionResponses, len(responses))

''' Computes and returns the mean, variance, and entropy 
    for a specific question based on the response distributions. 
    Makes use of numpy for the list arithmetic. '''
def computeMeanVarEnt(response, probs):
    negProb = [-i for i in probs]
    expectation = sum(np.multiply(response, probs))
    doubleExpectation = sum(np.multiply(np.square(response), probs))
    variance = np.subtract(doubleExpectation, np.square(expectation))
    entropy = sum(np.multiply(negProb, np.log2(probs)))
    return expectation, variance, entropy

''' Computes the Hellinger distance of each possible question pair. 
    Returns a dictionary of the values with an accompanying key 
    describing the two questions.'''
def calculateHellinger(questProbs, titles):
    hellinDict = {}
    divSqRtTwo = 1 / math.sqrt(2)
    for qA in range(26):
        for qB in range(qA + 1, 26):
            probA = questProbs.get(titles[qA])
            probB = questProbs.get(titles[qB])
            temp = np.sqrt(sum(np.square(np.subtract(np.sqrt(probA), np.sqrt(probB)))))
            hellin = np.multiply(np.sqrt(temp), divSqRtTwo)

            hellinKey = "%d, %d" % (qA + 1, qB + 1)
            hellinDict.update({hellinKey:hellin})
    return hellinDict


def createBarGraph(titles, distributions):
    nGroups = len(titles)
    colors = ['g', 'b', 'r', 'y', 'k']

    # bar graph setup
    plt.figure(figsize=(18, 3))
    index = np.arange(nGroups)
    barWidth = 0.15
    opacity = 0.8

    # creates graph data
    for rating in range(5):
        q = []
        counter = barWidth * (rating + 1)
        for keys in distributions:
            q.append(distributions.get(keys)[rating])
        plt.bar(index + counter, tuple(q), barWidth, alpha=opacity, 
                color=colors[rating], label=(rating + 1))   

    plt.xlabel('Questions')
    plt.ylabel('Distributions')
    plt.xticks(index + (barWidth * 3), tuple(titles))
    plt.legend()
    plt.show()


def createScatterPlot(hellinDict):
    x_axis = []
    for i in range(325):
        x_axis.append(i + 1)

    colors = (0, 0, 0)
    area = np.pi*3
 
    # Plot
    plt.scatter(x_axis, hellinDict.values(), s=area, c=colors, alpha=0.5)
    plt.title('Hellinger\'s Distance of Question Pairs')
    plt.xlabel('Pairs of Questions')
    plt.ylabel('Distance')
    plt.show()         

# Formats the data using the functions above and enters it into a dictionary for easy access. 
# will also create a dictionary that holds all the question totals
answer = list(range(1, 6))

for i in range(26):
    colLetter = chr(65 + i)
    questionName = "Q%d" % (i + 1)
    questionTitles.append(questionName)
    
    questionProb = distributionOfColumn(colLetter)
    questionMean, questionVar, questionEnt = computeMeanVarEnt(answer, questionProb)
 
    questionMeanVarEnt.update({questionName:[questionMean, questionVar, questionEnt]})          
    questionDistributions.update({questionName:questionProb})           

hellDict = calculateHellinger(questionDistributions, questionTitles)