import os
import numpy as np
import matplotlib.pyplot as plt
import win32com.client as win32
from scipy.optimize import minimize


def objectFun(x):
    hyApp = win32.Dispatch("HYSYS.Application")
    hysysSimulationCase = win32.GetObject("C:/ZY/05HYSYS/HYSYS_OPTIMIZATION_THF.hsc")
    Spreadsheet = hysysSimulationCase.Flowsheet.Operations.Item("SHEET")
    T100 = hysysSimulationCase.Flowsheet.Operations.Item("T-100")

    specTHF = T100.ColumnFlowsheet.Specifications.Item('THF_PURITY_SPEC')
    specToluene = T100.ColumnFlowsheet.Specifications.Item('TOLUENE_PURITY_SPEC')

    hysysSimulationCase.Solver.CanSolve = False

    specTHF.GoalValue = x[0]
    specToluene.GoalValue = x[1]

    hysysSimulationCase.Solver.CanSolve = True

    return -Spreadsheet.Cell(3, 7).CellValue



b1 = (0.9900,0.9990)
b2 = (0.9000,0.9900)

x0 = [0.9900,0.9000]
bnds = (b1,b2)

sol = minimize(objectFun,x0,method = 'SLSQP',bounds = bnds)

print(sol)


