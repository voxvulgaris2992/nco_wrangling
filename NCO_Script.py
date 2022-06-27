# The "rate limiting SKU": the minimum value to set the criterion is the calculation itself, not the initial minimum
# E.g. 4 < 6, but (4 / 2) still < (6 / 1)

import pandas as pd 
import math

df_SKUs = pd.read_excel('NCO_Script_Input.xlsx', usecols = [0])

df_qty = pd.read_excel('NCO_Script_Input.xlsx', usecols = [1])

SKUs = df_SKUs.values.tolist()

qty = df_qty.values.tolist()

Preset_4v1_SKUs = [["Fixed 120 cm"],["3-4 People"],["Bolia Armchair"],["Bolia Sofa"],["Task Chair"],["Visitor Chair"],["Desk-Side Storage 600 mm deep"],["Desk-Side Storage 800 mm deep "],["Tambour Desk Storage 600 mm"],["Tambour Desk Storage 800 mm "],["Bookcase"],["Tall Storage"]]

Preset_4v1_budget = []

i = 0

while i < len(Preset_4v1_SKUs):
	Preset_4v1_budget.append(qty[SKUs.index(Preset_4v1_SKUs[i])])
	i += 1

Preset_4v1_cons_marginal = [4,1,1,1,4,3,1,1,3,1,1,1]

Preset_4v1_multiplier = math.floor(qty[SKUs.index(["3-4 People"])][0]/Preset_4v1_cons_marginal[Preset_4v1_SKUs.index(["3-4 People"])])
#Amount of "Office for 4 v1 Presets" used

Preset_4v1_cons_actual = []

i = 0

while i < len(Preset_4v1_cons_marginal):
	Preset_4v1_cons_actual.append(Preset_4v1_cons_marginal[i]*Preset_4v1_multiplier)
	i += 1

i = 0

while i < len(Preset_4v1_budget):
	Preset_4v1_budget[i] = Preset_4v1_budget[i][0] - Preset_4v1_cons_actual[i]
	i += 1

i = 0 

while i < len(qty):
	qty[i] = qty[i][0]
	i += 1

i = 0

while i < len(Preset_4v1_SKUs):
	qty[SKUs.index(Preset_4v1_SKUs[i])] = Preset_4v1_budget[i]
	i += 1

Preset_4v2_SKUs = [["Fixed 120 cm"],["Bolia Armchair"],["Bolia Sofa"],["Bolia Coffee Table"],["Round Side Table"],["Task Chair"],["Visitor Chair"],["Desk-Side Storage 600 mm deep"],["Desk-Side Storage 800 mm deep "],["Tambour Desk Storage 600 mm"],["Tambour Desk Storage 800 mm "],["Bookcase"],["Tall Storage"]]

Preset_4v2_budget = []

i = 0

while i < len(Preset_4v2_SKUs):
	Preset_4v2_budget.append(qty[SKUs.index(Preset_4v2_SKUs[i])])
	i += 1

Preset_4v2_cons_marginal = [4,1,1,1,1,4,3,1,1,3,1,1,1]

Preset_4v2_multiplier = math.floor(qty[SKUs.index(["Bolia Coffee Table"])]/Preset_4v2_cons_marginal[Preset_4v2_SKUs.index(["Bolia Coffee Table"])])
#Amount of "Office for 4 v2 Presets" used

Preset_4v2_cons_actual = []

i = 0

while i < len(Preset_4v2_cons_marginal):
	Preset_4v2_cons_actual.append(Preset_4v2_cons_marginal[i]*Preset_4v2_multiplier)
	i += 1

i = 0

while i < len(Preset_4v2_budget):
	Preset_4v2_budget[i] = Preset_4v2_budget[i] - Preset_4v2_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_4v2_SKUs):
	qty[SKUs.index(Preset_4v2_SKUs[i])] = Preset_4v2_budget[i]
	i += 1

Preset_2_SKUs = [["Fixed 120 cm"],["Be Free Cube Seat"],["Be Free Small Cube"],["Task Chair"],["Visitor Chair"],["Pouffe"],["Desk-Side Storage 600 mm deep"],["Desk-Side Storage 800 mm deep "],["Tambour Desk Storage 600 mm"],["Tall Storage"]]

Preset_2_budget = []

i = 0

while i < len(Preset_2_SKUs):
	Preset_2_budget.append(qty[SKUs.index(Preset_2_SKUs[i])])
	i += 1

Preset_2_cons_marginal = [2,1,1,2,2,1,1,1,2,1]

Preset_2_multiplier = math.floor(qty[SKUs.index(["Be Free Cube Seat"])]/Preset_2_cons_marginal[Preset_2_SKUs.index(["Be Free Cube Seat"])])
#Amount of "Office for 2 Presets" used

Preset_2_cons_actual = []

i = 0

while i < len(Preset_2_cons_marginal):
	Preset_2_cons_actual.append(Preset_2_cons_marginal[i]*Preset_2_multiplier)
	i += 1

i = 0

while i < len(Preset_2_budget):
	Preset_2_budget[i] = Preset_2_budget[i] - Preset_2_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_2_SKUs):
	qty[SKUs.index(Preset_2_SKUs[i])] = Preset_2_budget[i]
	i += 1

Preset_100cm_SKUs = [["Fixed 100 cm"],["Task Chair"]]
Preset_100cm_budget = []

i = 0

while i < len(Preset_100cm_SKUs):
	Preset_100cm_budget.append(qty[SKUs.index(Preset_100cm_SKUs[i])])
	i += 1

Preset_100cm_cons_marginal = [1,1]

Preset_100cm_multiplier = math.floor(qty[SKUs.index(["Fixed 100 cm"])]/Preset_100cm_cons_marginal[Preset_100cm_SKUs.index(["Fixed 100 cm"])])
#Amount of "Office for 1 (100cm desks) Presets" used

Preset_100cm_cons_actual = []

i = 0

while i < len(Preset_100cm_cons_marginal):
	Preset_100cm_cons_actual.append(Preset_100cm_cons_marginal[i]*Preset_100cm_multiplier)
	i += 1

i = 0

while i < len(Preset_100cm_budget):
	Preset_100cm_budget[i] = Preset_100cm_budget[i] - Preset_100cm_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_100cm_SKUs):
	qty[SKUs.index(Preset_100cm_SKUs[i])] = Preset_100cm_budget[i]
	i += 1

Preset_140cm_SKUs = [["Fixed 140 cm"],["Task Chair"]]
Preset_140cm_budget = []

i = 0

while i < len(Preset_140cm_SKUs):
	Preset_140cm_budget.append(qty[SKUs.index(Preset_140cm_SKUs[i])])
	i += 1

Preset_140cm_cons_marginal = [1,1]

Preset_140cm_multiplier = math.floor(qty[SKUs.index(["Fixed 140 cm"])]/Preset_100cm_cons_marginal[Preset_140cm_SKUs.index(["Fixed 140 cm"])])
#Amount of "Office for 1 (140cm desks) Presets" used

Preset_140cm_cons_actual = []

i = 0

while i < len(Preset_140cm_cons_marginal):
	Preset_140cm_cons_actual.append(Preset_140cm_cons_marginal[i]*Preset_140cm_multiplier)
	i += 1

i = 0

while i < len(Preset_140cm_budget):
	Preset_140cm_budget[i] = Preset_140cm_budget[i] - Preset_140cm_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_140cm_SKUs):
	qty[SKUs.index(Preset_140cm_SKUs[i])] = Preset_140cm_budget[i]
	i += 1

Preset_160cm_SKUs = [["Fixed 160 cm"],["Task Chair"]]
Preset_160cm_budget = []

i = 0

while i < len(Preset_160cm_SKUs):
	Preset_160cm_budget.append(qty[SKUs.index(Preset_160cm_SKUs[i])])
	i += 1

Preset_160cm_cons_marginal = [1,1]

Preset_160cm_multiplier = math.floor(qty[SKUs.index(["Fixed 160 cm"])]/Preset_160cm_cons_marginal[Preset_160cm_SKUs.index(["Fixed 160 cm"])])
#Amount of "Office for 1 (160cm desks) Presets" used

Preset_160cm_cons_actual = []

i = 0

while i < len(Preset_160cm_cons_marginal):
	Preset_160cm_cons_actual.append(Preset_160cm_cons_marginal[i]*Preset_160cm_multiplier)
	i += 1

i = 0

while i < len(Preset_160cm_budget):
	Preset_160cm_budget[i] = Preset_160cm_budget[i] - Preset_160cm_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_160cm_SKUs):
	qty[SKUs.index(Preset_160cm_SKUs[i])] = Preset_160cm_budget[i]
	i += 1

Preset_120cm_SKUs = [["Fixed 120 cm"],["Task Chair"]]
Preset_120cm_budget = []

i = 0

while i < len(Preset_120cm_SKUs):
	Preset_120cm_budget.append(qty[SKUs.index(Preset_120cm_SKUs[i])])
	i += 1

Preset_120cm_cons_marginal = [1,1]

Preset_120cm_multiplier = math.floor(qty[SKUs.index(["Fixed 120 cm"])]/Preset_120cm_cons_marginal[Preset_120cm_SKUs.index(["Fixed 120 cm"])])
#Amount of "Office for 1 (120cm desks) Presets" used

Preset_120cm_cons_actual = []

i = 0

while i < len(Preset_120cm_cons_marginal):
	Preset_120cm_cons_actual.append(Preset_120cm_cons_marginal[i]*Preset_120cm_multiplier)
	i += 1

i = 0

while i < len(Preset_120cm_budget):
	Preset_120cm_budget[i] = Preset_120cm_budget[i] - Preset_120cm_cons_actual[i]
	i += 1

i = 0

while i < len(Preset_120cm_SKUs):
	qty[SKUs.index(Preset_120cm_SKUs[i])] = Preset_120cm_budget[i]
	i += 1

residual_SKUs = []

i = 0

while i < len(qty):
	if qty[i] != 0:
		residual_SKUs.append(SKUs[i][0])
	i += 1

residual_qty = []

i = 0

while i < len(qty):
	if qty[i] != 0:
		residual_qty.append(qty[i])
	i += 1

print("Preset Quantities:")
print("Office for 4 v1 = " + str(Preset_4v1_multiplier))
print("Office for 4 v2 = " + str(Preset_4v2_multiplier))
print("Office for 2 = " + str(Preset_2_multiplier))
print("Office for 1 (100cm fixed desks) = " + str(Preset_100cm_multiplier))
print("Office for 1 (140cm fixed desks) = " + str(Preset_140cm_multiplier))
print("Office for 1 (160cm fixed desks) = " + str(Preset_160cm_multiplier))
print("Office for 1 (120cm fixed desks) = " + str(Preset_120cm_multiplier))

print(" ")

print("Residual Furniture:")

i = 0

while i < len(residual_SKUs):
	print(residual_SKUs[i] + ": " + str(residual_qty[i]))
	i += 1























