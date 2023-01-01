#!/usr/bin/python3

import pandas as pd
from termcolor import colored

excelInputPath = "PythonExcelExamples\\example01.xlsx"

def start():
    print(colored(" Starting Example 1 ", "cyan"))
    print(colored("|------------------|", "cyan"))

    analysisDF = pd.read_excel(excelInputPath, sheet_name="TestPercent")

    # Call funciton we defined below to add 2 columns of percentage change
    analysisDF = createPercentColumns(analysisDF)

    # Write to Excel (*NO* formatting)
    #analysisDF.to_excel(excelInputPath, sheet_name="TestPercent", index=False)

    # Write to Excel (*WITH* formatting)
    
    writer = pd.ExcelWriter(excelInputPath, engine='xlsxwriter')
    analysisDF.to_excel(writer, sheet_name='TestPercent', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['TestPercent']
    red_format = workbook.add_format({'bg_color':'red'})
    green_format = workbook.add_format({'bg_color':'green'})

    worksheet.conditional_format('E2:E100', {'type': 'cell',
                                            'criteria': '<',
                                            'value':     '0',
                                            'format': red_format})

    worksheet.conditional_format('F2:F100', {'type': 'cell',
                                            'criteria': '>',
                                            'value':     '0',
                                            'format':  green_format})
    writer.save()
    # DF = dataframe
    #namesDF = pd.read_excel(excelInputPath, sheet_name="Sheet1")
    #print(colored("Columns: " + namesDF.columns, "cyan"))
    #print(f"Columns: {namesDF.columns}")
    #print(namesDF)

def createPercentColumns(analysisDF):
    # analysisDF = pd.read_excel(excelInputPath, sheet_name="TestPercent")
    print(analysisDF)

    # Create a place
    pctHList = []
    pctLList = []
    # Proccess the High and Low percent change for each row
    for i, row in analysisDF.iterrows():
        # Calculate percent
        pctHList.append((row["CLOSE"] - row["52WH"])/row["52WH"])
        pctLList.append((row["CLOSE"] - row["52WL"])/row["52WL"])

    analysisDF["%H"] = pctHList
    analysisDF["%L"] = pctLList

    return analysisDF


# Entry point into program
if __name__ == "__main__":
    start()
