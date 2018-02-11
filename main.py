import cyrtranslit
import openpyxl
import os
from lxml import etree
import consts
import numpy

def write_to_file(root, filename):
    if (not os.path.isdir('.\output')):
        os.mkdir('.\output')
    os.chdir('.\output')
    text_file = open(filename, "wb")
    text_file.write(etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8'))
    text_file.close()
    print('Data written to file')

def read_file(fileName,rangeEnd, columnName, columnYear, collabGraph, sumOfWorks, professorsDict):
    professors_papers = openpyxl.load_workbook(os.path.join(fileName))
    papersSheet = professors_papers.active
    print(fileName + ' file reading...')
    for i in range(2, rangeEnd):
        authors = cyrtranslit.to_latin(papersSheet[columnName + str(i)].value)
        try:
            year = int(papersSheet[columnYear + str(i)].value)
        except TypeError:
            year = 0
        except ValueError:
            year = 0
        if year < 2000 or year > 2016:
            continue
        removableChars = '{}"'
        for char in removableChars:
            authors = authors.replace(char,'')
        authors = authors.split(',')
        if ('Sanja Delčev' in authors):
            print(authors)
        authorsCopy = [i for i in authors]
        for author in authors:
            if not (author in professorsDict):
                authorsCopy.remove(author)
        authors = authorsCopy

        for i in range(0, len(authors)):
            sumOfWorks[professorsDict[authors[i]]['id']] += 1
            for j in range(i + 1, len(authors)):
                m = professorsDict[authors[i]]['id']
                n = professorsDict[authors[j]]['id']
                collabGraph[m][n] = collabGraph[m][n] + 1


def num_of_zeros(array):
    num = 0
    for item in array:
        if item == 0:
            num += 1
    return num

def average_coauthors(collabGraph):
    average_array = [0] * collabGraph.shape[0]
    index = 0
    for row in collabGraph:
        if len(row) - num_of_zeros(row) == 0:
            continue
        average_array[index] = (sum(row)/(len(row) - num_of_zeros(row)))
        index += 1
    return average_array

def array_sums(collabGraph):
    sums = [0] * collabGraph.shape[0]
    index = 0
    for row in collabGraph:
        sums[index] = sum(row)
        index += 1
    return sums
def collab_departments(collabGraph, professorsList):
    num_different = 0
    num_same = 0
    size = collabGraph.shape[0]
    for i in range(0, size):
        for j in range(0, size):
            if (professorsList[i]['department'] != professorsList[j]['department']) and (collabGraph[i][j] > 0):
                num_different += 1
            elif (professorsList[i]['department'] == professorsList[j]['department']) and (collabGraph[i][j] > 0):
                num_same += 1
    return {'same':num_same, 'different':num_different}

gexf = etree.Element('gexf', xmlns ='http://www.gexf.net/1.2draft', version='1.2')
graph = etree.SubElement(gexf, 'graph', mode='static', defaultedgetype='undirected')
#attributeNode = etree.SubElement(graph, 'attributes', class='node')
attributeNode = etree.SubElement(graph, 'attributes', mode='static')
attributeNode.set('class','node')
etree.SubElement(attributeNode, 'attribute', id= '0', title='ukupno', type='int')
etree.SubElement(attributeNode, 'attribute', id= '1', title='katedra', type='string')
attributeEdge = etree.SubElement(graph, 'attributes', mode='static')
attributeEdge.set('class', 'edge')
nodes = etree.SubElement(graph, 'nodes')

professors = openpyxl.load_workbook(os.path.join(consts.PROFESSORS_FILE_PATH))
professorsSheet = professors.active
print('Professors file reading...')
professorsDict = {}
professorsList = []

for i in range (2, 255):
    katedra = cyrtranslit.to_latin(professorsSheet['A' + str(i)].value)
    ime = cyrtranslit.to_latin(professorsSheet['D' + str(i)].value)
    professorsDict[ime] = {ime: katedra, 'id': i-2}
    professorsList.append({'name': ime, 'department': katedra})


collabGraph = numpy.zeros(shape = (253, 253))
sumOfWorks = [0] * collabGraph.shape[0]

read_file(consts.PROFESSORS_NATIVE_PATH, 844,'M', 'L', collabGraph, sumOfWorks, professorsDict)
read_file(consts.PROFESSORS_PAPERS_PATH, 2317, 'B', 'H', collabGraph, sumOfWorks, professorsDict)
read_file(consts.PROFESSORS_INTERNATIONAL_PATH, 887, 'M', 'L', collabGraph, sumOfWorks, professorsDict)
row_sums = array_sums(collabGraph)

for i in range(0, 253):
    node = etree.SubElement(nodes, 'node', id= str(i), label = professorsList[i]['name'])
    attvalues = etree.SubElement(node, 'attvalues')
    attvalueSum = etree.SubElement(attvalues,'attvalue', value=str(int(sumOfWorks[i])))
    attvalueSum.set('for','0')
    attvalueDepartment = etree.SubElement(attvalues,'attvalue',value=professorsList[i]['department'])
    attvalueDepartment.set('for','1')

edges = etree.SubElement(graph, 'edges')

index = 0;
for i in range(0,253):
    for j in range(0,253-i):
        edge = etree.SubElement(edges, 'edge', id = str(index), source= , target=)
        index += 1

average_coauthors(collabGraph)
print(collab_departments(collabGraph, professorsList))

write_to_file(gexf,'gephidata.gexf')

#print(listaProfesora)