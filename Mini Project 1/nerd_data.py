from openpyxl import load_workbook                 
from scipy.stats import rv_discrete
import numpy as np
import matplotlib.pyplot as plt
import math


# opening .xlsx file
wb = load_workbook('mini-project1-data.xlsx')                                                                            
ws = wb[wb.sheetnames[0]]                                          
ws = wb.active                                              

# Global Variables
questionDistributions = {}
questionMeanVarEnt = {}
questionTitles = []

'''
Parses the accompanying .xlsx file extracting each response. The responses 
    are then separated by type (1 - 5). The rates and total responses are 
    returned.

Parameters:
    columnLabel - the name of the column being parsed
'''
def dataByColumn(columnLabel):
    questionResponses = [0 for i in range(5)]
    responses = []

    # parse .xlsx file and add to list
    for row in range(2, ws.max_row + 1):
        for column in columnLabel:
            cellName = "{}{}".format(column, row)
            if ws[cellName].value != 0:
                responses.append(ws[cellName].value)        # omit the 0s

    # populate list with rates per response
    for i in range(5):
        questionResponses[i] = responses.count(i + 1)
    
    return questionResponses, len(responses)

'''

'''
def distributionOfQuestions(numResponse, total):
    prob = np.divide(numResponse,total)
    return prob

'''
Takes the question titles and distributions as parameters.
The data is formatted and a bar graph is created. 
Uses the matplotlib to draw the graph.
'''
def createBarGraph(titles, distributions):
    nGroups = len(titles)
    colors = ['g', 'b', 'r', 'y', 'k']

    # bar graph setup
    plt.figure(figsize=(18,3))
    index = np.arange(nGroups)
    barWidth = 0.15
    opacity = 0.8

    # creates graph data
    for rating in range(5):
        q = []
        counter = barWidth * (rating + 1)
        for key in distributions:
            q.append(distributions.get(key)[rating])
        plt.bar(index + counter, tuple(q), barWidth, alpha=opacity, color=colors[rating], label=(rating + 1))   

    plt.xlabel('Questions')
    plt.ylabel('Distributions')
    plt.xticks(index + (barWidth * 3), tuple(titles))
    plt.legend()
    plt.show()

'''
Function computes and returns the mean, variance, and entropy 
    of a set of distributions. Makes use of the scipy.stats
    library.

Parameters: 
    rates - a list of sums of all specific responses for a question
    prob - a list of the probabilities of all responses for a question
Returns: mean, variance, entropy
'''
def computeMeanVarEnt(rates, prob):
    negProb = [-i for i in prob]
    rv = rv_discrete(values=(rates, prob))

    ent = np.multiply(negProb, np.log2(prob))
    return rv.mean(), rv.var(), sum(ent)

# TODO: loop through the questions
# TODO: extract the prob lists, and do calculations
# TODO: store in a dictionary {question numbers as string:Hellinger calculation}
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

def createScatterPlot(hellinDict):
    x_axis = []
    for i in range(325):
        x_axis.append(i + 1)

    colors = (0,0,0)
    area = np.pi*3
 
    # Plot
    plt.scatter(x_axis, hellinDict.values(), s=area, c=colors, alpha=0.5)
    plt.title('Hellinger\'s Distance of Question Pairs')
    plt.xlabel('Pairs of Questions')
    plt.ylabel('Distance')
    plt.show()         

# Formats the data using the functions above and enters it into a dictionary for easy access. 
# will also create a dictionary that holds all the question totals
for i in range(26):
    colLetter = chr(65 + i)
    questionName = "Q%d" % (i + 1)
    questionTitles.append(questionName)
    
    responseRates, numOfResponses = dataByColumn(colLetter) 
    questionProb = distributionOfQuestions(responseRates, numOfResponses)
    questionMean, questionVar, questionEnt = computeMeanVarEnt(responseRates, questionProb)
 
    questionMeanVarEnt.update({questionName:[questionMean,questionVar, questionEnt]})          
    questionDistributions.update({questionName:questionProb})           

hellDict = calculateHellinger(questionDistributions, questionTitles)

minHellinValue = min(hellDict.values())
minHellinKey = [key for key in hellDict if hellDict[key] == minHellinValue]
#print(minHellinKey)
#print(minHellinValue)
#print(hellDict)

#print(questionMeanVarEnt)
#createBarGraph(questionTitles, questionDistributions)
#createScatterPlot(hellDict)
createBarGraph(questionTitles, questionDistributions)