import pandas as pd
from decimal import Decimal


def goal_seek(f, a, b, tol):
    """
    Finds a root of the function f in the interval [a, b] to within a tolerance tol.
    """
    x=a
    a=Decimal(a)
    b=Decimal(b)
    x= Decimal(x)

    while abs(f(a)) > tol:
        x = Decimal((Decimal(a) + Decimal(b))) / Decimal(2)
        if f(x) == 0:
            return x
        elif Decimal(f(a)) * Decimal(f(x)) < 0:
            b = x
        else:
            a = x

    return round(((a + b) / 2),15)

def f(x):
    # Calculate the values of F column and returns the value of the function in F102.
    r1 = 0
    r2 = 1

    temp_col = [0]
    
    for i in range(rowNum):
        aRow = df.at[r2, 'a']
        cRow = df.at[r2, 'c']
        eRow = df.at[r2, 'e']
       
        
        temp_col.append(Decimal(temp_col[r1]) * Decimal((1 + eRow)) + (Decimal(aRow) * Decimal(x) - Decimal(cRow)))
       
        r1 += 1
        r2 += 1
# Since excel sheet has year 0 our df columns has length 101 if your original df does not have year zerro
#change below code with this
#   df['g'] = temp_col.pop(0)
#   return df.at[99,'g']

    df['g'] = temp_col
    return df.at[100,'g']

# Number of rows in data frame
rowNum = 100


# Load the data frame
df = pd.read_excel('test.xlsx')

# Define columns and rows used in the calculations
ByChangingCellColumn = df.columns[8]
ByChangingCellRow = 4
calcOfFunctionsCol = df.columns[6]

# Calculate the desired value of I6
a = df.at[ByChangingCellRow, ByChangingCellColumn]
b = a + 1

result = goal_seek(f, a, b, tol=1e-8)
print("By changing cell: " + str(result))
