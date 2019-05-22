# -*- coding: utf-8 -*-

import csv, re
from openpyxl import Workbook, load_workbook
from collections import namedtuple

filename     = 'WorldBirdList-TiF-taxonomy-June-2018.xlsx'
treeSheet    = 'Tree (TiF)'
speciesSheet = 'List (TiF)'

Aidx=ord("A")
def colIdx(colName):
    if len(colName) == 2:
        return 26*(ord(colName[0])-Aidx+1) + ord(colName[1])-Aidx
    else:
        return ord(colName)-Aidx

treeCols = [(colIdx(c[0]), c[1]) for c in [('B', 'clade'),('D', 'clade'),('F', 'clade'),('H', 'clade'),('J', 'parvclass'),('K', 'clade'),('N', 'clade'),('R', 'superorder'),('Z', 'order'),('AE', 'suborder'),('AI', 'infraorder'),('AM', 'parvorder'),('AY', 'superfamily'),('BP', 'family'),('BR', 'subfamily'),('BT', 'tribe')]]
treeEnglishCol = colIdx('BW')
treeGenusCol   = colIdx('CI')
treeFirstRow = 2
treeMaxRow   = 1800

class Taxon:
    id=None
    pid=None
    col=None
    rank=None
    name=None
    eng=None

    def __init__(self, col, rank, name, eng):
        self.id = name
        self.pid = None
        self.col = col
        self.rank = rank
        self.name = name
        self.eng = eng

    def __str__(self):
        return "{:12} {:20} [id={}, pid={}]  {}".format(self.rank, self.name, self.id, self.pid, self.eng)

synMatcher = re.compile('^(.+) *\[= *(.+) *] *')

parents = []    

def readTreeRow(ws, row):
    for idx, col in enumerate(treeCols):
        val = row[col[0]].value
        #print(idx, col, val)
        if val:
            common = row[treeEnglishCol].value
            return Taxon(col=idx, rank=col[1], name=val, eng=common)
    return None

def write(t):
    nout.write("{},{},{}\n".format(t.id, t.rank, t.name))
    tout.write("{},{},{},accepted\n".format(t.id, t.pid if t.pid else "", t.id))
    if t.eng:
        vout.write("{},eng,\"{}\"\n".format(t.id, t.eng))

def parseTree(ws):
    print("Parse tree")
    print(ws.calculate_dimension())
    #ws.reset_dimensions()
    for idx, row in enumerate(ws.rows):
        if idx < treeFirstRow:
            continue
        t = readTreeRow(ws, row)
        if t:
            while(parents and parents[-1].col >= t.col):
                parents.pop()
            t.pid = parents[-1].id if parents else None
            m = synMatcher.search(t.name)
            if m:
                print(m)
            print("{:4} {}".format(idx, t))
            write(t)
            parents.append(t)


print("Open spreadsheet")
wb = load_workbook(filename = filename, read_only=True)
with open('name.csv', 'w', newline='') as nout:
    nout.write("ID,rank,scientificName\n")
    with open('taxon.csv', 'w', newline='') as tout:
        tout.write("ID,parentID,nameID,status\n")
        with open('vernacularname.csv', 'w', newline='') as vout:
            vout.write("taxonID,language,name\n")
            parseTree(wb[treeSheet])
