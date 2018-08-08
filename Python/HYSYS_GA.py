# Import libraries
import win32com.client as win32
import matplotlib.pyplot as plt

import os
import sys

from mpl_toolkits.mplot3d import Axes3D


from gaft import GAEngine
from gaft.components import GAIndividual
from gaft.components import GAPopulation
from gaft.operators import RouletteWheelSelection
from gaft.operators import UniformCrossover
from gaft.operators import FlipBitMutation

# Built-in best fitness analysis.
from gaft.analysis.fitness_store import FitnessStoreAnalysis
from gaft.analysis.console_output import ConsoleOutputAnalysis

str_THF_LB = float(sys.argv[1])
str_THF_HB = float(sys.argv[2])
str_TOL_LB = float(sys.argv[3])
str_TOL_HB = float(sys.argv[4])
str_P = int(sys.argv[5])
str_PC = float(sys.argv[6])
str_PM = float(sys.argv[7])
str_ng = int(sys.argv[8])

#str_THF_LB = 0.99
#str_THF_HB = 0.999
#str_TOL_LB = 0.90
#str_TOL_HB = 0.99
#str_P = 20
#str_PC = 0.8
#str_PM = 0.1
#str_ng = 2

# Define population.
indv_template = GAIndividual(ranges=[(str_THF_LB,str_THF_HB), (str_TOL_LB,str_TOL_HB)],
                             encoding='binary',
                             eps=0.001)
population = GAPopulation(indv_template=indv_template, size=str_P).init()

# Create genetic operators.
selection = RouletteWheelSelection()
crossover = UniformCrossover(pc=str_PC, pe=0.5)
mutation = FlipBitMutation(pm=str_PM)

# Create genetic algorithm engine.
# Here we pass all built-in analysis to engine constructor.
engine = GAEngine(population=population, selection=selection,
                  crossover=crossover, mutation=mutation,
                  analysis=[ConsoleOutputAnalysis, FitnessStoreAnalysis])

# Define fitness function.
@engine.fitness_register
def fitness(indv):
    # Active HYSYS Model
    hyApp_Fitness = win32.Dispatch("HYSYS.Application")
    hyApp_Fitness.Visible = True
    hysysSimulationCase_Fitness = hyApp_Fitness.ActiveDocument

    # Objective Function
    Spreadsheet = hysysSimulationCase_Fitness.Flowsheet.Operations.Item("SHEET")
    T100 = hysysSimulationCase_Fitness.Flowsheet.Operations.Item("T-100")

    specTHF = T100.ColumnFlowsheet.Specifications.Item('THF_PURITY_SPEC')
    specToluene = T100.ColumnFlowsheet.Specifications.Item('TOLUENE_PURITY_SPEC')

    hysysSimulationCase_Fitness.Solver.CanSolve = False

    specTHF.GoalValue, specToluene.GoalValue = indv.variants

    hysysSimulationCase_Fitness.Solver.CanSolve = True

    ax.scatter(specTHF.GoalValue, specToluene.GoalValue, Spreadsheet.Cell(3, 7).CellValue)
    #plt.pause(0.0001)

    return Spreadsheet.Cell(3, 7).CellValue


if '__main__' == __name__:
    # Open the Hysys Model
    hyApp = win32.Dispatch("HYSYS.Application")
    hyApp.Visible = True
    hySimulationCase = hyApp.ActiveDocument

    fig = plt.figure(figsize=(16,9))
    ax = fig.add_subplot(111,projection = '3d')

    ax.set_xlabel('THF_PURITY_SPEC')
    ax.set_ylabel('TOLUENE_PURITY_SPEC')
    ax.set_zlabel('Profit')
    #ax.set_title('')

    #plt.ion()
    #plt.show()

    # Run Genetic Algorithm
    ga_best = engine.run(ng=str_ng)

    #ax.scatter(ga_best[0][0], ga_best[0][1], ga_best[1], c='r', marker='^', s=160, edgecolor='0.3',alpha=0.2)
    ax.scatter(ga_best[0][0], ga_best[0][1], ga_best[1], c='r', marker='^', s=160)

    plt.savefig('C:/ZY/03Output/Profit.png')

    print(ga_best)