# Business Decision Support System with Excel Solver

This repository contains several classic operations research problems solved using **Excel Solver**, a powerful optimization tool included in Microsoft Excel.

## Table of Contents

- [Overview](#overview)
- [Problems Solved](#problems-solved)
- [Excel Solver Usage](#excel-solver-usage)
- [File Structure](#file-structure)
- [How to Use](#how-to-use)
- [License](#license)

---

## Overview

The objective of this project is to model and solve optimization problems related to logistics and resource allocation. These include:

- **Transportation problems** (minimizing shipping costs)
- **Knapsack problems** (maximizing value within a capacity constraint)
- **Facility location problems** (minimizing operational and delivery costs)

All problems are formulated and solved using **Microsoft Excel's Solver Add-in**.

---

## Problems Solved

### 1. Transportation Problems
Sheets: `Transport1 (3p)`, `Transport2 (3p)`, `Transport3 (9p)`

- **Objective**: Minimize the total cost of transporting goods from multiple factories to several customers.
- **Constraints**: Supply constraints at factories and demand requirements at customer locations.

### 2. Knapsack Problem
Sheet: `Knapsack (9p)`

- **Scenario**: A fuel truck with multiple compartments must be loaded with gas types.
- **Objective**: Minimize loss from unfulfilled gas demand.
- **Constraints**: Compartment sizes and binary loading decisions.

### 3. Facility Location Problem
Sheet: `Facility (9p)`

- **Objective**: Determine which facilities to operate and how to route shipments to minimize fixed + transport costs.
- **Constraints**: Each warehouse must be served; plants have limited capacity.

---

## Excel Solver Usage

Solver is used to find optimal solutions under constraints. Here's how it was applied:

- **Objective Function**: Defined as a total cost/loss cell calculated using formulas.
- **Decision Variables**: Represented as cells corresponding to quantities shipped, items selected, or facilities chosen.
- **Constraints**:
  - Demand/supply balances
  - Capacity limits
  - Binary restrictions for yes/no decisions (e.g., plant open or not)

### Steps to Use Solver:

1. Go to `Data` → `Solver`
2. Set the objective cell (e.g., total cost)
3. Choose `Min` for minimization
4. Define decision variable cells
5. Add constraints (e.g., supply ≤ capacity)
6. Choose `Simplex LP` or `GRG Nonlinear` depending on problem type
7. Click `Solve`

> **Note**: Some problems (like Knapsack) use binary or integer constraints, so `Solver Options → Assume Linear Model` must be unchecked and the `Simplex LP` method is usually selected.

---

## File Structure

- `Business Decision Support System with Solver.xlsx` — Excel file with all problems and Solver models
  - `Summary` — Overview
  - `Transport*` — Transportation problems (3 and 9 points)
  - `Knapsack` — Knapsack problem with partial loading
  - `Facility` — Facility location optimization

---

## How to Use

1. Open `Business Decision Support System with Solver.xlsx` in Microsoft Excel.
2. Navigate to the relevant sheet.
3. Use `Solver` from the `Data` tab to re-run or tweak the model.
4. Modify input values (e.g., costs, capacities) to experiment with different scenarios.

---

## License

This project is shared for educational purposes and learning. Please credit the author when using it for your own work or extensions.

