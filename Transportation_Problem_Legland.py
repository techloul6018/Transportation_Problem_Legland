import pandas as pd
import numpy as np
from itertools import product

def read_excel(file_path):
    data = pd.read_excel(file_path, header=None)
    costs = data.iloc[:3, :3].values
    supply = data.iloc[:3, 3].values
    demand = data.iloc[3, :3].values
    return costs, supply, demand

def northwest_corner(costs, supply, demand):
    rows, cols = costs.shape
    allocation = np.zeros((rows, cols))
    i, j = 0, 0
    while i < rows and j < cols:
        alloc = min(supply[i], demand[j])
        allocation[i, j] = alloc
        supply[i] -= alloc
        demand[j] -= alloc
        if supply[i] == 0:
            i += 1
        elif demand[j] == 0:
            j += 1
    return allocation

def minimum_cost_method(costs, supply, demand):
    rows, cols = costs.shape
    allocation = np.zeros((rows, cols))
    cost_indices = sorted(product(range(rows), range(cols)), key=lambda x: costs[x])
    for i, j in cost_indices:
        if supply[i] == 0 or demand[j] == 0:
            continue
        alloc = min(supply[i], demand[j])
        allocation[i, j] = alloc
        supply[i] -= alloc
        demand[j] -= alloc
    return allocation

def vogels_method(costs, supply, demand):
    rows, cols = costs.shape
    allocation = np.zeros((rows, cols))
    
    supply_left = supply.copy()
    demand_left = demand.copy()
    supply_points = list(range(rows))
    demand_points = list(range(cols))

    while supply_points and demand_points:
        penalties = []

        for i in supply_points:
            row_costs = [costs[i][j] for j in demand_points]
            if len(row_costs) >= 2:
                sorted_costs = sorted(row_costs)
                penalties.append((sorted_costs[1] - sorted_costs[0], i, 'row'))
            elif len(row_costs) == 1:
                penalties.append((row_costs[0], i, 'row'))

        for j in demand_points:
            col_costs = [costs[i][j] for i in supply_points]
            if len(col_costs) >= 2:
                sorted_costs = sorted(col_costs)
                penalties.append((sorted_costs[1] - sorted_costs[0], j, 'col'))
            elif len(col_costs) == 1:
                penalties.append((col_costs[0], j, 'col'))

        if not penalties:
            break

        penalty, index,