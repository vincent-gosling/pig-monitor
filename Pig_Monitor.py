import pandas as pd
import xlsxwriter
import os
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta


# input the data from the input sheets
def input_excel(df, month):
    # input Excel spreadsheet data
    inputs = pd.read_excel(f"figures 20{month[3:]}\\{month} inputs.xlsx")
    inputs = [inputs.columns.tolist()] + inputs.values.tolist()

    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15

    for i in range(1, 80):
        df[i].insert(0, 0)

    # boars inputs
    df[1][0] = inputs[4][B]
    df[3][0] = inputs[5][B] + inputs[6][B]
    df[4][0] = inputs[7][B]
    df[5][0] = inputs[9][B] + inputs[11][B]
    df[6][0] = inputs[10][B]

    # sows and gilts inputs
    df[7][0] = inputs[4][C]
    df[8][0] = inputs[4][D]
    df[11][0] = inputs[5][D] + inputs[6][D]
    df[12][0] = inputs[7][D]
    df[13][0] = inputs[9][C]
    df[14][0] = inputs[10][C]
    df[15][0] = inputs[11][C]
    df[16][0] = inputs[9][D]
    df[17][0] = inputs[10][D]
    df[18][0] = inputs[11][D]
    df[19][0] = inputs[13][D]
    df[20][0] = inputs[13][C]
    df[21][0] = inputs[14][C]

    # piglets inputs
    df[22][0] = inputs[4][H]
    df[25][0] = inputs[15][C]
    df[26][0] = inputs[5][H]
    df[27][0] = inputs[6][H]
    df[28][0] = inputs[7][H]
    df[29][0] = inputs[9][H]
    df[30][0] = inputs[16][C]
    df[31][0] = inputs[8][H]
    df[32][0] = round(df[26][0] / df[25][0], 7)
    df[33][0] = round(df[31][0] / df[30][0], 7)
    df[34][0] = round(inputs[18][E] / df[31][0], 7)
    df[35][0] = inputs[7][E] / df[31][0]

    # sow and piglet feed inputs
    df[38][0] = inputs[7][O]
    df[37][0] = inputs[7][M] - df[38][0] + inputs[7][K]
    df[40][0] = inputs[10][O]
    df[39][0] = inputs[10][M] - df[40][0] + inputs[10][K]
    df[41][0] = df[37][0] + df[39][0]
    df[42][0] = df[38][0] + df[40][0]
    df[43][0] = inputs[7][N] - inputs[7][P] + inputs[7][L]
    df[44][0] = inputs[10][N] - inputs[10][P] + inputs[10][L]
    df[45][0] = df[43][0] + df[44][0]

    # finishers inputs
    df[46][0] = inputs[4][E]
    df[49][0] = inputs[5][E]
    df[50][0] = round(inputs[18][E] / df[49][0], 7)
    df[51][0] = inputs[7][E] / df[49][0]
    df[52][0] = inputs[19][E]
    try:
        df[53][0] = inputs[20][E] / df[52][0]
    except:
        df[53][0] = 0
    try:
        df[54][0] = inputs[21][E] / df[52][0]
    except:
        df[54][0] = 0
    df[55][0] = inputs[11][E]

    # finisher feed inputs
    df[57][0] = inputs[20][O]
    df[56][0] = inputs[20][M] - df[57][0] + inputs[20][K]
    df[58][0] = inputs[20][N] - inputs[20][P] + inputs[20][L]

    # sales inputs
    df[59][0] = inputs[9][E]
    df[60][0] = inputs[10][E]
    df[61][0] = round(df[60][0] / df[59][0], 7)
    df[62][0] = inputs[22][E]
    df[63][0] = inputs[23][E]
    df[64][0] = inputs[24][E]
    df[65][0] = round(df[62][0] / df[59][0], 7)
    df[66][0] = round(df[63][0] / df[59][0], 7)
    df[67][0] = inputs[25][E]

    # evaluation inputs
    df[2][0] = inputs[13][H]
    df[9][0] = inputs[14][H]
    df[10][0] = inputs[15][H]
    df[23][0] = inputs[16][H]
    df[24][0] = inputs[17][H]
    df[36][0] = inputs[18][H]
    df[47][0] = inputs[19][H]
    df[48][0] = inputs[20][H]

    # costs inputs
    df[68][0] = inputs[22][H]
    df[69][0] = inputs[23][H]
    df[70][0] = inputs[24][H]
    df[71][0] = inputs[25][H]
    df[72][0] = inputs[22][J]
    df[73][0] = inputs[23][J]
    df[74][0] = inputs[24][J]
    df[75][0] = inputs[25][J]
    df[76][0] = inputs[23][N]
    df[77][0] = inputs[24][N]
    df[78][0] = inputs[25][N]
    df[79][0] = inputs[7][L] + inputs[10][L] + inputs[20][L]

    return df


# programs calculations
def output_calcs(df, data):
    # performance calculations
    sows_lagged, days = [0, 0, 0], [0, 0, 0]
    for i in range(3):
        n = 3 * 2 ** i
        for j in range(0, n):  # Days in period
            days[i] += calendar.monthrange(2000 + int(str(df[0][j])[3:]), int(str(df[0][j])[:2]))[1]
        sows_lagged[i] = sum(df[7][4:n+4]) / n  # average number of sows and served gilts (lagged)
        data[3][3][i+1] = sum(df[8][:n]) / n  # average number of unserved gilts
        data[3][4][i+1] = sum(df[7][:n]) / n  # average number of sows and served gilts
        data[2][3][i+1] = data[3][4][i+1]
        data[3][5][i+1] = data[3][4][i+1] / sum(df[1][:n]) * n  # sow/boar ratio
        data[3][6][i+1] = sum(df[11][:n]) / (data[2][3][i+1] / 400 * 2 ** i)  # sow replacement rate (%)
        data[3][7][i+1] = sum(df[13][:n]) / (data[2][3][i+1] / 400 * 2 ** i)  # sow sales (%)
        data[3][8][i+1][0] = sum(df[15][:n]) / (data[2][3][i+1] / 400 * 2 ** i)  # sow deaths (%)
        data[3][10][i+1] = sum(df[19][:n] + df[20][:n])  # number of 1st services
        data[3][11][i+1] = sum(df[21][:n])  # number of repeat services
        data[3][12][i+1] = (data[3][10][i+1] - data[3][11][i+1]) * 100 / data[3][10][i+1]  # conception to first service (%)
        data[3][13][i+1][0] = sum(df[25][:n]) * 100 / sum(df[19][4:n+4] + df[20][4:n+4] + df[21][4:n+4])  # farrowing rate (%)
        data[2][4][i+1] = data[3][13][i+1][0]
        data[3][14][i+1][0] = sum(df[25][:n]) / sum(df[7][4:n+4]) * (2 ** (2 - i)) * n  # litters per sow per year
        data[2][5][i+1][0] = data[3][14][i+1][0]
        data[3][25][i+1] = sum(df[36][: n]) / n  # average age at weaning
        data[3][15][i+1][0] = 365 / data[3][14][i+1][0] - 115 - data[3][25][i+1]  # empty days per sow per litter
        data[3][17][i+1] = sum(df[25][:n])  # number of litters born
        data[3][18][i+1] = sum(df[26][:n] + df[27][:n] + df[28][:n])  # Total piglets born
        data[3][19][i+1][0] = sum(df[26][:n]) / data[3][17][i+1]  # Piglets born alive per litter
        data[2][6][i+1] = data[3][19][i+1][0]
        data[3][20][i+1] = sum(df[27][:n]) / data[3][17][i+1]  # Piglets born dead per litter
        data[3][21][i+1] = sum(df[28][:n]) / data[3][17][i+1]  # Piglets born mummified per litter
        data[3][22][i+1] = sum(df[29][:n]) * 100 / data[3][19][i+1][0] / data[3][17][i+1]  # piglet mortality to weaning (%)
        data[2][8][i+1] = data[3][22][i+1]
        data[3][23][i+1][0] = sum(df[31][:n]) / sum(df[30][:n])  # piglets weaned per litter
        data[2][7][i+1][0] = data[3][23][i+1][0]
        data[3][24][i+1][0] = sum(df[31][:n]) * (2 ** (2 - i)) / sows_lagged[i]  # piglets reared per sow per year
        data[2][9][i+1][0] = data[3][24][i+1][0]
        data[3][26][i+1][0] = sum(df[34][:n]) / n  # average weight at weaning
        data[3][28][i+1] = sum(df[37][:n]) * 2 ** (2 - i) / (data[3][3][i+1] + data[3][4][i+1] + data[3][4][i+1] / data[3][5][i+1])  # Sow feed used/sow,gilt,boar/year (T)
        data[3][29][i+1] = sum(df[39][:n]) * 1000 / sum(df[31][:n])  # Piglet feed used/piglet to weaning (kg)
        data[3][30][i+1] = (sum(df[39][:n] + df[37][:n])) * 1000 / sum(df[31][:n])  # Total feed used per piglet weaned (kg)
        data[3][31][i+1][0] = sum(df[31][:n]) / sum(df[37][:n])  # Piglets reared per T of sow feed
        data[4][3][i+1] = sum(df[46][:n]) / n  # Average pigs on unit
        data[2][11][i+1] = data[4][3][i+1]
        data[4][4][i+1] = sum(df[49][:n])  # Number of pigs trans in
        data[4][5][i+1] = sum(df[59][:n])  # Number of pigs sold (inc cas)
        data[4][6][i+1] = sum(df[52][:n])  # Number of pigs trans out
        data[4][7][i+1] = sum(df[50][:n]) / n  # Average liveweight per pig in (kg)
        data[2][12][i+1] = data[4][7][i+1]
        data[4][8][i+1] = sum([df[63][j] + df[52][j] * df[53][j] for j in range(n)]) / sum(df[59][:n] + df[52][:n])  # Average liveweight per pig out (kg)
        data[2][13][i+1] = data[4][8][i+1]
        data[4][9][i+1] = sum(df[65][:n]) / n  # Average deadweight per pig sold (inc cas) (kg)
        data[4][10][i+1][0] = sum(df[67][:n]) / n  # Average P2 (mm)
        data[2][17][i+1][0] = data[4][10][i+1][0]
        data[4][12][i+1] = sum(df[56][:n]) / (data[4][5][i+1] + data[4][6][i+1]) * 1000  # Feed used per pig out (kg)
        data[4][13][i+1][0] = sum(df[56][:n])  # total feed used (T)
        data[4][15][i+1][0] = sum(df[56][:n]) / data[4][3][1] * 1000 / days[i]  # Daily feed intake per pig (kg)
        data[2][14][i+1][0] = data[4][15][i+1][0]
        data[4][16][i+1][0] = sum([df[46][j] * df[47][j] + df[52][j] * df[53][j] + df[63][j] - df[46][j+1] * df[47][j+1] - df[34][j] * df[31][j] for j in range(n)]) / data[4][3][i+1] / days[i] * 1000  # Daily liveweight gain (g)
        data[2][15][i+1][0] = data[4][16][i+1][0]
        data[4][17][i+1][0] = data[4][13][i+1][0] * 1000 / (data[4][16][i+1][0] * days[i] / 1000 * data[4][3][i+1])  # Liveweight FCR
        data[2][16][i+1][0] = data[4][17][i+1][0]
        data[4][18][i+1] = 100 / sum(df[66][:n]) * data[4][9][i+1] * n  # KO %
        data[4][19][i+1][0] = data[4][17][i+1][0] / data[4][18][i+1] * 100  # Carcase FCR
        data[4][20][i+1] = data[4][16][i+1][0] / 100 * data[4][18][i+1]  # Daily carcase gain (g)
        data[4][21][i+1] = 0.533 * data[4][16][i+1][0]  # Daily lean gain (g)
        data[2][18][i+1][0] = data[4][21][i+1]
        data[4][22][i+1] = sum(df[55][:n]) * 200 / sum(df[49][:n] + df[52][:n] + df[59][:n] + df[55][:n])  # Mortality (%)
        data[4][23][i+1][0] = (data[4][8][i+1] - data[4][7][i+1]) * 1000 / data[4][16][i+1][0]  # days in unit
        n = 4 * i + 2 ** i
        data[4][33][i+1][0] = sum(df[59][:n])  # Number sold deadweight
        data[4][34][i+1][0] = sum(df[62][:n]) / sum(df[59][:n])  # Average carcase weight (kg)
        data[4][35][i+1] = sum(df[60][:n]) / sum(df[62][:n]) * 100  # Average carcase p/kg
        data[4][36][i+1] = sum([df[64][j] * df[59][j] for j in range(n)]) / sum(df[59][:n])  # % Grade 1
        data[4][37][i+1] = sum(df[65][:n]) / sum(df[66][:n]) * 100  # KO %
        data[4][38][i+1] = sum([df[67][j] * df[59][j] for j in range(n)]) / sum(df[59][:n])  # Avg P2
        data[4][41][i+1] = sum(df[60][:n])  # Total sales value (£)
        data[4][42][i+1][0] = sum(df[60][:n]) / sum(df[59][:n])  # Average value per pig (£)

    # finance calculations
    per_pig = [0, 0, 0]
    for i in range(6):
        n = 4 * int(i / 2) + 2 ** int(i / 2)
        per_pig[int(i / 2)] = sum(df[31][:n])  # piglets weaned
        data[5][4][i+1][0] = sum(df[6][:n] + df[14][:n])  # value of sows/boars sold
        data[5][5][i+1][0] = sum([df[31][j] * df[35][j] for j in range(n)])  # value of piglets transferred
        data[5][6][i+1][0] = data[5][4][i+1][0] + data[5][5][i+1][0]  # total breeders output
        data[5][8][i+1][0] = sum(df[4][:n] + df[12][:n])  # value of gilts and boars bought
        data[5][9][i+1][0] = df[1][0] * df[2][0] - df[1][n] * df[2][n] + sum([df[j][0] * df[j + 2][0] - df[j][n] * df[j + 2][n] for j in [7, 8, 22]])  # valuation change
        data[5][10][i+1][0] = data[5][6][i+1][0] - data[5][8][i+1][0] + data[5][9][i+1][0]  # livestock output
        data[5][12][i+1][0] = sum(df[43][:n])  # cost of sow feed
        data[5][13][i+1][0] = sum(df[44][:n])  # cost of piglet feed
        data[5][14][i+1][0] = data[5][12][i+1][0] + data[5][13][i+1][0]  # total feed cost
        data[5][16][i+1][0] = data[5][10][i+1][0] - data[5][14][i+1][0]  # margin over feed cost
    for i in [4, 5, 6, 8, 9, 10, 12, 13, 14, 16]:
        for j in [1, 3, 5]:
            data[5][i][j+1][0] /= per_pig[int((j - 1) / 2)]
    for i in range(6):
        n = 4 * int(i / 2) + 2 ** int(i / 2)
        per_pig[int(i / 2)] = sum(df[52][:n] + df[59][:n])  # pigs sold
        data[6][4][i+1][0] = sum(df[60][:n])  # value of pigs sold
        data[6][5][i+1][0] = sum([df[52][j] * df[54][j] for j in range(n)])  # value of pigs transferred out
        data[6][6][i+1][0] = data[6][4][i+1][0] + data[6][5][i+1][0]  # total output
        data[6][8][i+1][0] = sum([df[49][j] * df[51][j] for j in range(n)])  # value of pigs transferred in
        data[6][9][i+1][0] = df[46][0] * df[48][0] - df[46][n] * df[48][n]  # valuation change
        data[6][10][i+1][0] = data[6][6][i+1][0] - data[6][8][i+1][0] + data[6][9][i+1][0]  # livestock output
        data[6][12][i+1][0] = sum(df[58][:n])  # feed cost
        data[6][13][1+2*int(i / 2)][0] = data[6][12][1+2*int(i / 2)][0] / data[6][10][1+2*int(i / 2)][0] * 100  # Feed Cost per £100 Output (£)
        data[6][14][1+2*int(i / 2)][0] = data[6][12][1+2*int(i / 2)][0] / sum([df[46][j] * df[47][j] + df[52][j] * df[53][j] + df[63][j] - df[46][j + 1] * df[47][j + 1] - df[34][j] * df[31][j] for j in range(n)]) * 100  # Feed Cost per kg Liveweight Gain (p)
        data[6][15][1+2*int(i / 2)][0] = data[6][14][1+2*int(i / 2)][0] * 100 / data[4][18][1+int(i / 2)]  # Feed Cost per kg Carcase Gain (p)
        data[6][17][i+1][0] = data[6][10][i+1][0] - data[6][12][i+1][0]  # Margin Over Feed Cost (£)
    for i in [4, 5, 6, 8, 9, 10, 12, 17]:
        for j in [1, 3, 5]:
            data[6][i][j + 1][0] /= per_pig[int((j - 1) / 2)]
    for i in range(6):
        n = 4 * int(i / 2) + 2 ** int(i / 2)
        per_pig[int(i / 2)] = sum(df[59][:n])  # pigs sold
        data[7][4][i+1][0] = data[5][10][i+1][0]  # livestock output breeding herd
        data[7][5][i+1][0] = data[6][10][i+1][0]  # Livestock Output Feeding Herd
        data[7][6][i+1][0] = data[7][4][i+1][0] + data[7][5][i+1][0]  # Livestock Output Combined Herd
        data[7][8][i+1][0] = data[6][12][i+1][0] + data[5][14][i+1][0]  # Feed Cost
        data[7][10][i+1][0] = data[6][17][i+1][0] + data[5][16][i+1][0]  # Margin Over Feed Cost
        for j in range(26, 34):  # Variable Costs
            data[7][j-14][i+1][0] = sum(df[j + 42][:n])
        data[7][20][i+1][0] = sum([data[7][j][i+1][0] for j in range(12, 20)])  # Total Variable costs
        data[7][21][i+1][0] = data[7][10][i+1][0] - data[7][20][i+1][0]  # Gross Margin
        data[7][23][i+1][0] = sum(df[76][:n])  # Labour
        data[7][24][i+1][0] = sum(df[77][:n])  # Buildings
        data[7][25][i+1][0] = sum(df[78][:n])  # Other fixed costs
        data[7][26][i+1][0] = data[7][25][i+1][0] + data[7][24][i+1][0] + data[7][23][i+1][0]  # Total Fixed costs
        data[7][27][i+1][0] = data[7][21][i+1][0] - data[7][26][i+1][0]  # Net Margin
        data[7][28][1+2*int(i / 2)][0] = sum(df[41][:n] + df[56][:n]) / sum(df[63][:n]) * 1000  # Whole Herd FCR
        data[7][29][1+2*int(i / 2)][0] = (data[7][26][1+2*int(i / 2)][0] + data[7][20][1+2*int(i / 2)][0] + data[7][8][1+2*int(i / 2)][0]) / sum(df[62][:n]) * 100  # Total Cost per kg Carcase (p)
        data[7][30][1+2*int(i / 2)][0] = data[4][35][1+int(i / 2)]  # Average Price/kg Carcase (p)
    for i in [4, 5, 6, 8, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27]:
        for j in [1, 3, 5]:
            data[7][i][j+1][0] = data[7][i][j][0] / per_pig[int((j - 1) / 2)]

    # cashflow calculations
    for i in range(11, -1, -1):
        data[8][2][i+1][0] = df[6][i] + df[14][i] + df[60][i]
        data[8][3][i+1][0] = df[79][i]
        for j in range(6, 14):
            data[8][j][i+1] = df[j+62][i]
        data[8][14][i+1][0] = sum(row[i+1] for row in data[8][6:14])
        data[8][17][i+1] = df[76][i]
        data[8][18][i+1] = df[77][i]
        data[8][19][i+1] = df[78][i]
        data[8][20][i+1][0] = sum(row[i+1] for row in data[8][17:20])
        data[8][22][i+1][0] = data[8][2][i+1][0] - data[8][3][i+1][0] - data[8][14][i+1][0] - data[8][20][i+1][0]
        data[8][23][i+1][0] = sum(row[0] for row in data[8][22][i+1:13])

    for i in [2, 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 17, 18, 19, 20]:
        if isinstance(data[8][i][13], list):
            data[8][i][13] = sum(row[0] for row in data[8][i][1:13])
        else:
            data[8][i][13] = sum(data[8][i][1:13])


# full data sheets program
def pig_monitor(i):
    # open previous months datafile
    old_df = []

    month = datetime.strptime(i, "%m-%y")
    last_month = month - relativedelta(months=1)
    last_month = last_month.strftime("%m-%y")

    with open(f"data files\\{last_month}.txt", 'r') as file:
        for j in file.readlines():
            old_df.append(j.split())

    # correct to create this month's datafile
    for j in range(80):
        old_df[j].pop()
    old_df[0].insert(0, i.replace('-', '/'))

    df = input_excel(old_df, i)

    for j in range(1, len(df)):
        df[j] = [float(element) for element in df[j]]

    # write to this month's datafile
    with open(f"data files\\{i}.txt", 'w') as myfile:
        pass

        for j in range(16):
            myfile.write((14 - len(str(df[0][j]))) * ' ' + str(df[0][j]))
        myfile.write('\n')

        for j in range(1, 80):
            for k in range(16):
                myfile.write((14 - len(str(round(df[j][k], 6)))) * ' ' + str(round(df[j][k], 6)))
            myfile.write('\n')

    # prepare files to add new output sheet
    if os.path.exists(f"figures 20{i[3:]}\\{i} figures.xlsx"):
        os.remove(f"figures 20{i[3:]}\\{i} figures.xlsx")

    # Create workbook and worksheet
    workbook = xlsxwriter.Workbook(f"figures 20{i[3:]}\\{i} figures.xlsx")
    worksheets = []

    for j in ['Input1', 'Input2', 'KPF', 'Breeder', 'Finisher', 'Breed £', 'Fin £', 'Combined £', '£']:
        worksheets.append(workbook.add_worksheet(j))

    # Add formats
    reg = workbook.add_format({'bg_color': 'white'})
    reg_ = workbook.add_format({'bg_color': 'white'})
    reg_.set_bottom()
    sub = workbook.add_format({'bg_color': 'white', 'bold': True, 'underline': True})
    special = workbook.add_format({'bg_color': 'white', 'font_color': 'red', 'bold': 'True'})
    special_ = workbook.add_format({'bg_color': 'white', 'font_color': 'red', 'bold': 'True'})
    special_.set_bottom()
    date = workbook.add_format({'bg_color': 'white', 'align': 'right'})
    date.set_bottom()
    head = workbook.add_format({'bg_color': 'white', 'bold': True, 'underline': True, 'align': 'center', 'font_size': 20})

    # specific formats for finance sheets
    month = workbook.add_format({'bg_color': 'white', 'align': 'center'})
    month.set_left()
    month.set_right()
    month_ = workbook.add_format({'bg_color': 'white', 'align': 'center'})
    month_.set_bottom()
    month_.set_left()
    month_.set_right()
    special2 = workbook.add_format({'bg_color': 'white', 'font_color': 'red', 'bold': 'True'})
    special2.set_left()
    special2.set_right()
    special2_ = workbook.add_format({'bg_color': 'white', 'font_color': 'red', 'bold': 'True'})
    special2_.set_bottom()
    special2_.set_left()
    special2_.set_right()
    reg2 = workbook.add_format({'bg_color': 'white'})
    reg2.set_left()
    reg2.set_right()
    reg2_ = workbook.add_format({'bg_color': 'white'})
    reg2_.set_bottom()
    reg2_.set_left()
    reg2_.set_right()

    # data
    data = [0, 0, 0, 0, 0, 0, 0, 0, 0]
    data[0] = [[None, None, None, None, None, None, None, None, None, None, None, None, None],
               [[None, reg_], ['05/25', date], ['04/25', date], ['03/25', date], ['02/25', date], ['01/25', date], ['12/24', date], ['11/24', date], ['10/24', date], ['09/24', date], ['08/24', date], ['07/24', date], ['06/24', date]],
               ['Boars                            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of boars               ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Boars bought/transferred in      ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Value of boars in (£)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Boars sold/died                  ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Receipts from boars sold (£)     ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               ['Sows                             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Maiden gilts                     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of sows (£)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of maiden gilts (£)    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gilts bought/transferred in      ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Value of gilts in (£)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sows sold                        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Receipts from sows sold (£)      ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sows died                        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gilts sold                       ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Receipts from gilts sold (£)     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gilts died                       ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gilts served 1st time            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sows served 1st time             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Sow/gilt repeat services         ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               ['Piglets                          ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg weight of piglets            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of piglets (£)         ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Litters born                     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglets born alive               ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglets born dead                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglets born mummified           ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglet deaths                    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Litters weaned                   ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglets weaned                   ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg piglets alive/litter         ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg piglets weaned/litter        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg weight piglets weaned (kg)   ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of piglets weaned (£)  ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Avg weaning age (days)           ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               ['Sow feed used (T)                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sow feed stock (T)               ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglet feed used (T)             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglet feed stock (T)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sow & Piglet feed used (T)       ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sow & Piglet feed stock (T)      ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Sow feed used (£)                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Piglet feed used (£)             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Sow & Piglet feed cost (£)       ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]]]
    data[1] = [[None, None, None, None, None, None, None, None, None, None, None, None, None],
               [[None, reg_], ['05/25', date], ['04/25', date], ['03/25', date], ['02/25', date], ['01/25', date], ['12/24', date], ['11/24', date], ['10/24', date], ['09/24', date], ['08/24', date], ['07/24', date], ['06/24', date]],
               ['Number of pigs                   ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg weight of pigs (kg)          ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value of pigs (£)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Number of pigs trans in          ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg weight pigs trans in (kg)    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value pigs trans in (£)      ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Number of pigs trans out         ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg weight pigs trans out (kg)   ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value pigs trans out (£)     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Number of pig deaths             ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               ['Weight of feed used (T)          ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Weight of feed at end (T)        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Cost of feed used (£)            ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               ['Pigs sold                        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Receipts (£)                     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg value (£)                    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Total carcase weight (kg)        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Total liveweight (kg)            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['% pigs in grade 1                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg carcase weight (kg)          ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Avg liveweight (kg)              ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Avg P2 (mm)                      ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]],
               [None, None, None, None, None, None, None, None, None, None, None, None, None],
               [None, None, None, None, None, None, None, None, None, None, None, None, None],
               [[None, reg_], ['05/25', date], ['04/25', date], ['03/25', date], ['02/25', date], ['01/25', date], ['12/24', date], ['11/24', date], ['10/24', date], ['09/24', date], ['08/24', date], ['07/24', date], ['06/24', date]],
               ['Feed medication                  ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Vet and medicines                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gas and  electricity             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Water                            ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Straw and bedding                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Transport                        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['AI charges                       ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Miscellaneous                    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Labour                           ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Buildings                        ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Other fixed costs                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['combined feed costs              ', reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_], [0, reg_]]]
    data[2] = [[None, None, None, None],
               [[None, reg_], ['Last 3 months', date], ['Last 6 months', date], ['Last 12 months', date]],
               [['BREEDERS', sub], None, None, None],
               ['Avg number of sows and served gilts            ', 0, 0, 0],
               ['Farrowing rate (%)                             ', 0, 0, 0],
               [['Litters per sow per year                       ', special], [0, special], [0, special], [0, special]],
               ['Piglets born alive per litter                  ', 0, 0, 0],
               [['Piglets weaned per litter                      ', special], [0, special], [0, special], [0, special]],
               ['Piglet mortality to weaning (%)                ', 0, 0, 0],
               [['Piglets reared per sow per year                ', special_], [0, special_], [0, special_], [0, special_]],
               [['COMBINED FEEDING HERD', sub], None, None, None],
               ['Average pigs on unit                           ', 0, 0, 0],
               ['Average liveweight per pig in (kg)             ', 0, 0, 0],
               ['Average liveweight per pig out (kg)            ', 0, 0, 0],
               [['Daily feed intake per pig (kg)                 ', special], [0, special], [0, special], [0, special]],
               [['Daily liveweight gain (g)                      ', special], [0, special], [0, special], [0, special]],
               [['Liveweight FCR                                 ', special], [0, special], [0, special], [0, special]],
               [['Average P2 (mm)                                ', special], [0, special], [0, special], [0, special]],
               [['Daily lean gain (g)                            ', reg_], [0, reg_], [0, reg_], [0, reg_]]]
    data[3] = [[None, None, None, None],
               [[None, reg_], ['Last 3 months', date], ['Last 6 months', date], ['Last 12 months', date]],
               [['Herd Structure', sub], None, None, None],
               ['Average number of unserved gilts               ', 0, 0, 0],
               ['Average number of sows and served gilts        ', 0, 0, 0],
               ['Sow/boar ratio                                 ', 0, 0, 0],
               ['Sow replacement rate (%)                       ', 0, 0, 0],
               ['Sow sales (%)                                  ', 0, 0, 0],
               [['Sow deaths (%)                                 ', reg_], [0, reg_], [0, reg_], [0, reg_]],
               [['Service Details', sub], None, None, None],
               ['Number of 1st services                         ', 0, 0, 0],
               ['Number of repeat services                      ', 0, 0, 0],
               ['% Conception to 1st service                    ', 0, 0, 0],
               [['Farrowing rate (%)                             ', special], [0, special], [0, special], [0, special]],
               [['Litters per sow per year                       ', special], [0, special], [0, special], [0, special]],
               [['Empty days per sow per litter                  ', reg_], [0, reg_], [0, reg_], [0, reg_]],
               [['Litter Details', sub], None, None, None],
               ['Number of litters born                         ', 0, 0, 0],
               ['Total piglets born                             ', 0, 0, 0],
               [['Piglets born alive per litter                  ', special], [0, special], [0, special], [0, special]],
               ['Piglets born dead per litter                   ', 0, 0, 0],
               ['Piglets born mummified per litter              ', 0, 0, 0],
               ['Piglet mortality to weaning (%)                ', 0, 0, 0],
               [['Piglets weaned per litter                      ', special], [0, special], [0, special], [0, special]],
               [['Piglets reared per sow per year                ', special], [0, special], [0, special], [0, special]],
               ['Average age at weaning                         ', 0, 0, 0],
               [['Average weight at weaning (kg)                 ', reg_], [0, reg_], [0, reg_], [0, reg_]],
               [['Feed Analysis', sub], None, None, None],
               ['Sow feed used/sow,gilt,boar/year (T)           ', 0, 0, 0],
               ['Piglet feed used/piglet to weaning (kg)        ', 0, 0, 0],
               ['Total feed used per piglet weaned (kg)         ', 0, 0, 0],
               [['Piglets reared per T of sow feed               ', reg_], [0, reg_], [0, reg_], [0, reg_]]]
    data[4] = [[None, None, None, None],
               [[None, reg_], ['Last 3 months', date], ['Last 6 months', date], ['Last 12 months', date]],
               [['Pig Movements', sub], None, None, None],
               ['Average pigs on unit                           ', 0, 0, 0],
               ['Number of pigs trans in                        ', 0, 0, 0],
               ['Number of pigs sold (inc cas)                  ', 0, 0, 0],
               ['Number of pigs trans out                       ', 0, 0, 0],
               ['Average liveweight per pig in (kg)             ', 0, 0, 0],
               ['Average liveweight per pig out (kg)            ', 0, 0, 0],
               ['Average deadweight per pig sold (inc cas) (kg) ', 0, 0, 0],
               [['Average P2 (mm)                                ', special_], [0, special_], [0, special_], [0, special_]],
               [['Feed Analysis', sub], None, None, None],
               ['Feed used per pig out (kg)                     ', 0, 0, 0],
               [['Total feed used (T)                            ', reg_], [0, reg_], [0, reg_], [0, reg_]],
               [['Performance', sub], None, None, None],
               [['Daily feed intake per pig (kg)                 ', special], [0, special], [0, special], [0, special]],
               [['Daily liveweight gain (g)                      ', special], [0, special], [0, special], [0, special]],
               [['Liveweight FCR                                 ', special], [0, special], [0, special], [0, special]],
               ['KO %                                           ', 0, 0, 0],
               [['Carcase FCR                                    ', special], [0, special], [0, special], [0, special]],
               ['Daily carcase gain (g)                         ', 0, 0, 0],
               ['Daily lean gain (g)                            ', 0, 0, 0],
               ['Mortality (%)                                  ', 0, 0, 0],
               [['Days in unit                                   ', reg_], [0, reg_], [0, reg_], [0, reg_]],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [None, None, None, None],
               [[None, reg_], ['This month', date], ['Last 6 months', date], ['Last 12 months', date]],
               [['Number sold deadweight                         ', special], [0, special], [0, special], [0, special]],
               [['Average carcase weight (kg)                    ', special], [0, special], [0, special], [0, special]],
               ['Average carcase p/kg                           ', 0, 0, 0],
               ['% Grade 1                                      ', 0, 0, 0],
               ['KO %                                           ', 0, 0, 0],
               ['Avg P2                                         ', 0, 0, 0],
               ['Avg Lean Meat %                                ', 0, 0, 0],
               ['Carcase Lean % (calc)                          ', 0, 0, 0],
               ['Total sales value (£)                          ', 0, 0, 0],
               [['Average value per pig (£)                      ', reg_], [0, reg_], [0, reg_], [0, reg_]]]
    data[5] = [[None, None, None, None, None, None, None],
               [None, None, None, None, None, None, None],
               [[None, reg_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_]],
               [['Outputs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Value of Sows & Boars sold (£)        ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Value of Piglets sold/transferred (£) ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Total Output (£)                      ', reg_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_]],
               [['Inputs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Value of Gilts & Boars transferred in ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Valuation Change (£)                  ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Livestock Output (£)                  ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]],
               [['Feed Costs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Cost of Sow Feed (£)                  ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Cost of Piglet Feed (£)               ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Total Cost of Feed (£)                ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]],
               [['Margin', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               [['Margin Over Feed Cost (£)             ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]]]
    data[6] = [[None, None, None, None, None, None, None],
               [None, None, None, None, None, None, None],
               [[None, reg_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_]],
               [['Outputs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Value of pigs sold (£)                ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Value of pigs transferred out (£)     ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Total Output (£)                      ', reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_]],
               [['Inputs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Value of pigs transferred in (£)      ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Valuation Change (£)                  ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Livestock Output (£)                  ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]],
               [['Feed Costs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               [['Feed Cost (£)                         ', special], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2]],
               ['Feed Cost per £100 Output (£)         ', [0, reg2], [None, reg2], [0, reg2], [None, reg2], [0, reg2], [None, reg2]],
               ['Feed Cost per kg Liveweight Gain (p)  ', [0, reg2], [None, reg2], [0, reg2], [None, reg2], [0, reg2], [None, reg2]],
               [['Feed Cost per kg Carcase Gain (p)     ', reg_], [0, reg2_], [None, reg2_], [0, reg2_], [None, reg2_], [0, reg2_], [None, reg2_]],
               [['Margin', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               [['Margin Over Feed Cost (£)             ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]]]
    data[7] = [[None, None, None, None, None, None, None],
               [None, None, None, None, None, None, None],
               [[None, reg_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_], ['HERD', month_], ['PER PIG', month_]],
               [['OUTPUTS (£)', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Livestock Output Breeding Herd        ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Livestock Output Feeding Herd         ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Livestock Output Combined Herd        ', reg_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_]],
               [['FEED COSTS (£)', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               [['Feed Cost                             ', reg_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_], [0, reg2_]],
               [['MARGINS (£)', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               [['Margin Over Feed Cost                 ', special], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2]],
               [['Variable Costs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Feed medication                       ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Vet & medicines                       ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Gas & electricity                     ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Water                                 ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Straw & bedding                       ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Transport                             ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['AI charges                            ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Miscellaneous                         ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Total Variable costs                  ', special], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2]],
               [['Gross Margin                          ', special], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2]],
               [['Fixed Costs', sub], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2], [None, reg2]],
               ['Labour                                ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Buildings                             ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               ['Other fixed costs                     ', [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2], [0, reg2]],
               [['Total Fixed costs                     ', special], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2], [0, special2]],
               [['Net Margin                            ', special_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_], [0, special2_]],
               ['Whole Herd FCR                        ', [0, reg2], [None, reg2], [0, reg2], [None, reg2], [0, reg2], [None, reg2]],
               ['Total Cost per kg Carcase (p)         ', [0, reg2], [None, reg2], [0, reg2], [None, reg2], [0, reg], [None, reg2]],
               [['Average Price/kg Carcase (p)          ', reg_], [0, reg2_], [None, reg2_], [0, reg2_], [None, reg2_], [0, reg2_], [None, reg2_]]]
    data[8] = [[None, None, None, None, None, None, None, None, None, None, None, None, None, None],
               [[None, reg_], ['05/25', date], ['04/25', date], ['03/25', date], ['02/25', date], ['01/25', date], ['12/24', date], ['11/24', date], ['10/24', date], ['09/24', date], ['08/24', date], ['07/24', date], ['06/24', date], ['12-mths', date]],
               [['Net Output (£)                ', special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special]],
               [['Total value Feed purchased (£)', special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special]],
               [None, None, None, None, None, None, None, None, None, None, None, None, None, None],
               [['Variable Costs', sub], None, None, None, None, None, None, None, None, None, None, None, None, None],
               ['Feed medication (£)           ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Vet & medicines (£)           ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Gas & electricity (£)         ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Water (£)                     ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Straw & bedding (£)           ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Transport (£)                 ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['AI charges (£)                ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Miscellaneous (£)             ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Total Variable costs (£)      ', special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special]],
               [None, None, None, None, None, None, None, None, None, None, None, None, None, None],
               [['Fixed Costs', sub], None, None, None, None, None, None, None, None, None, None, None, None, None],
               ['Labour (£)                    ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Buildings (£)                 ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               ['Other fixed costs (£)         ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
               [['Total Fixed costs (£)         ', special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special], [0, special]],
               [[None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_], [None, reg_]],
               [['Cashflow for month (£)        ', special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [None, special_]],
               [['Cumulative cashflow (£)       ', special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [0, special_], [None, special_]]]

    # output calculations
    for j in [0, 1, 8]:
        for k in range(1, 13):
            data[j][1][k][0] = df[0][k-1]

    for j in range(1, 13):
        data[1][26][j][0] = df[0][j-1]

    for j in range(2, 47):
        for k in range(1, 13):
            if isinstance(data[0][j][k], list):
                data[0][j][k][0] = df[j-1][k-1]
            else:
                data[0][j][k] = df[j-1][k-1]

    for j in range(2, 24):
        for k in range(1, 13):
            if isinstance(data[1][j][k], list):
                data[1][j][k][0] = df[j+44][k-1]
            else:
                data[1][j][k] = df[j+44][k-1]

    for j in range(27, 39):
        for k in range(1, 13):
            if isinstance(data[1][j][k], list):
                data[1][j][k][0] = df[j+41][k-1]
            else:
                data[1][j][k] = df[j+41][k-1]

    output_calcs(df, data)

    # Write data with formats
    for j in range(9):
        indexes = [[str(chr(65+col)) + str(row) for col in range(len(data[j][0]))] for row in range(1, len(data[j]) + 1)]

        # set printing defaults
        worksheets[j].set_landscape()
        worksheets[j].fit_to_pages(1, 0)

        for k in range(len(data[j])):
            for l in range(len(data[j][0])):
                if isinstance(data[j][k][l], list):
                    if isinstance(data[j][k][l][0], float):
                        worksheets[j].write(indexes[k][l], round(data[j][k][l][0], 2), data[j][k][l][1])
                    else:
                        worksheets[j].write(indexes[k][l], data[j][k][l][0], data[j][k][l][1])
                else:
                    if isinstance(data[j][k][l], float):
                        worksheets[j].write(indexes[k][l], round(data[j][k][l], 2), reg)
                    else:
                        worksheets[j].write(indexes[k][l], data[j][k][l], reg)

    # Adjust column width
    for j in range(9):
        worksheets[j].set_column('A:A', 40)

    for j in [2, 3, 4]:
        worksheets[j].set_column('B:D', 33)
    for j in [5, 6, 7]:
        worksheets[j].set_column('B:G', 16)

    worksheets[0].set_column('B:M', 14)
    worksheets[1].set_column('B:M', 14)

    # Adjust row height
    for j in range(9):
        for k in range(len(data[j])):
            worksheets[j].set_row(k, 18)

    # add headers
    worksheets[0].merge_range('A1:M1', 'BREEDERS INPUT & CALCULATED DATA', head)
    worksheets[0].set_row(0, 26)
    worksheets[1].merge_range('A1:M1', 'FINISHERS INPUT & CALCULATED DATA', head)
    worksheets[1].set_row(0, 26)
    worksheets[1].merge_range('A26:M26', 'COSTS INPUT DATA', head)
    worksheets[1].set_row(25, 26)
    worksheets[2].merge_range('A1:D1', 'KEY PERFORMANCE FACTORS', head)
    worksheets[2].set_row(0, 26)
    worksheets[3].merge_range('A1:D1', 'BREEDERS PERFORMANCE', head)
    worksheets[3].set_row(0, 26)
    worksheets[4].merge_range('A1:D1', 'FINISHERS PERFORMANCE', head)
    worksheets[4].set_row(0, 26)
    worksheets[4].merge_range('A32:D32', 'FINISHERS : CARCASE DATA', head)
    worksheets[4].set_row(31, 26)
    worksheets[5].merge_range('A1:G1', 'BREEDERS : FINANCIAL PERFORMANCE', head)
    worksheets[5].set_row(0, 26)
    worksheets[6].merge_range('A1:G1', 'FINISHERS : FINANCIAL PERFORMANCE', head)
    worksheets[6].set_row(0, 26)
    worksheets[7].merge_range('A1:G1', 'COMBINED HERD FINANCIAL PERFORMANCE', head)
    worksheets[7].set_row(0, 26)
    worksheets[8].merge_range('A1:N1', 'COMBINED HERD CASHFLOW', head)
    worksheets[8].set_row(0, 26)

    for j in [5, 6, 7]:
        worksheets[j].merge_range('B2:C2', 'This month', month)
        worksheets[j].merge_range('D2:E2', 'Last 6 months', month)
        worksheets[j].merge_range('F2:G2', 'Last 12 months', month)

    # Set the default zoom level
    zoom = [75, 75, 115, 115, 115, 115, 115, 115, 105]
    for j in range(9):
        worksheets[j].set_zoom(zoom[j])

    # Save workbook
    workbook.close()
