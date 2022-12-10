import csv
import json
from os import listdir, mkdir
from os.path import exists
import sys
import re

# Specify grasscutter resource ExcelBinOutput path here
OLD_EXCEL = 'D:\\Genshin Impart\\Grasscutter3.3\\resources\\ExcelBinOutput\\'
TSV = './txt/'
TEMP_OUTPUT_DIR = './_ExcelBinOutput/'
OUTPUT_DIR = './ExcelBinOutput/'

if not exists(TEMP_OUTPUT_DIR):
    print('Missing _ExcelBinOutput.. creating folder')
    mkdir(TEMP_OUTPUT_DIR)

if not exists(OUTPUT_DIR):
    print('Missing ExcelBinOutput.. creating folder')
    mkdir(OUTPUT_DIR)

if not exists(TSV):
    print('TSV folder missing, generating...\nMake sure to put the .txt files inside it to generate excels and run again')
    mkdir(TSV)
    sys.exit(1)

for i in listdir(OLD_EXCEL):
    if(exists(TSV + i.replace('ExcelConfig', '').replace('.json', '.txt'))):
        print(
            f"Generating {i} from {i.replace('ExcelConfig','').replace('.json','.txt')}")
        dataDict = []
        with open(TSV + i.replace('ExcelConfig', '').replace('.json', '.txt'), encoding='utf-8') as csvReader:
            reader = csv.DictReader(csvReader, delimiter="\t")

            ogExcel = json.load(open(OLD_EXCEL+i, 'r', encoding='utf-8'))

            try:
                for row in reader:
                    obj = {}
                    for data in row:
                        if(row[data].__contains__(";")):
                            array = []
                            for j in row[data].split(";"):
                                try:
                                    array.append(int(j))
                                except:
                                    try:
                                        array.append(float(row[data]))
                                    except:
                                        array.append(j)
                            row[data] = array
                        elif(row[data].__contains__(",")):
                            array = []
                            for j in row[data].split(","):
                                try:
                                    array.append(int(j))
                                except:
                                    try:
                                        array.append(float(row[data]))
                                    except:
                                        array.append(j)
                            row[data] = array

                        m = re.match(r'\[(.*)\]', data)
                        if(m != None):
                            m = re.match(
                                r'\[(.*?)\]([1-9]?)(.*?)([1-9]{0,1}$)', data)
                            if(m.group(2) != ''):
                                if(m.group(1) not in obj):
                                    obj[m.group(1)] = [{}]
                                elif(len(obj[m.group(1)]) < int(m.group(2))):
                                    obj[m.group(1)].append({})

                                if(m.group(4) != ''):
                                    if(m.group(1) not in obj[m.group(1)][int(m.group(2)) - 1]):
                                        obj[m.group(1)][int(m.group(2)) -
                                                        1][m.group(3)] = []
                                    try:
                                        obj[m.group(1)][int(m.group(2)) -
                                                        1][m.group(3)].append(int(row[data]))
                                    except:
                                        try:
                                            obj[m.group(1)][int(m.group(2)) -
                                                            1][m.group(3)].append(
                                                float(row[data]))
                                        except:
                                            obj[m.group(1)][int(m.group(2)) -
                                                            1][m.group(3)].append(row[data])

                                else:
                                    obj[m.group(1)][int(m.group(2)) -
                                                    1][m.group(3)] = row[data]

                            else:
                                if(m.group(1) not in obj):
                                    obj[m.group(1)] = {}
                                if(m.group(4) != ''):
                                    if(m.group(3) not in obj[m.group(1)]):
                                        obj[m.group(1)][m.group(3)] = []

                                    try:
                                        obj[m.group(1)][m.group(3)].append(
                                            int(row[data]))
                                    except:
                                        try:
                                            obj[m.group(1)][m.group(3)].append(
                                                float(row[data]))
                                        except:
                                            obj[m.group(1)][m.group(
                                                3)].append(row[data])
                                else:
                                    obj[m.group(1)][m.group(3)] = row[data]
                        else:
                            m = re.match(r'(.*?)([1-9]{0,1}$)', data)
                            if(m.group(2) != ''):
                                if(m.group(1) not in obj):
                                    obj[m.group(1)] = []
                                try:
                                    obj[m.group(1)].append(int(row[data]))
                                except:
                                    try:
                                        obj[m.group(1)].append(
                                            float(row[data]))
                                    except:
                                        obj[m.group(1)].append(row[data])
                            else:
                                obj[data] = row[data]
                    dataDict.append(obj)
            except:
                print()

        json.dump(dataDict, open(TEMP_OUTPUT_DIR + i, 'w',
                  encoding='utf-8'), indent=2, ensure_ascii=False)


print('Loading nameTranslation.json')
with open('./nameTranslation.json', 'r', encoding='utf-8') as ntFile:
    ntDict = json.load(ntFile)

for f in listdir(TEMP_OUTPUT_DIR):
    with open(TEMP_OUTPUT_DIR+f, 'r', encoding='utf-8') as file:
        print(f"Applying nameTranslation for {f}")
        rep = dict((re.escape(k), v) for k, v in ntDict.items())
        pattern = re.compile("|".join(rep.keys()))
        result = pattern.sub(lambda m: rep[re.escape(m.group(0))], file.read())
        with open(OUTPUT_DIR + f, 'w', encoding='utf-8') as file:
            file.write(result)
