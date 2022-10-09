import pandas as pd
import xlwings as xw

# =================================================================
# Pandas DataFrame Converter (Part 1)
# =================================================================

df = pd.read_csv("../summer.csv")

wb = xw.Book()
sheet = wb.sheets[0]

sheet["A1"].value = df

sheet["A1"].options(pd.DataFrame, expand="table").value
# range index as floats! (reading column A as an index)

sheet["B1"].options(pd.DataFrame, expand="table", index=False).value
# range index as integers

sheet["A1"].options(pd.DataFrame, expand="table", index=False).value
# create a new range index and read the first column as a column

sheet["A1"].options(pd.DataFrame, expand="table",
                    index=False, header=False).value
# header=False >>> creates a new header (range format)

df = sheet["A1"].options(pd.DataFrame, expand="table",
                         index=False, header=False).value
df.columns = ["Year", "City", "Sport", "Discipline", "Athlete",
              "Country", "Gender", "Event", "Medal"]
df

wb.close()


# =================================================================
# Pandas DataFrame Converter (Part 2)
# =================================================================

wb = xw.Book()
sheet = wb.sheets[0]

df = pd.read_csv("../summer.csv")
df

sheet["A1"].options(index=False).value = df
sheet["A1"].options(pd.DataFrame, expand="table", index=False).value
sheet["A1"].expand().clear_contents()
sheet["A1"].options(index=False, header=False).value = df

# Excursus: Autofit

sheet["E1"].autofit()  # autofit a single cell
sheet["D1:E10"].columns.autofit()  # autofit columns in range object
sheet["A:I"].autofit()  # autofit full columns
sheet.cells.autofit()  # this works too.. just like VBA


# =================================================================
# Data Science Application:
# Inspect and Manipulate DataFrames in Excel
# =================================================================

df = pd.read_csv("../summer.csv")

# drop columns
df.drop(columns=["Discipline", "Event"], inplace=True)
df

# reindex columns
df = df.reindex(columns=["Athlete", "Medal", "Year", "City",
                         "Sport", "Country", "Gender"])
df

# rename columns
df.rename(columns={"Athlete": "Name", "City": "Host_City",
                   "Country": "Nationality"}, inplace=True)
df

df = pd.read_csv("../summer.csv")

wb = xw.Book()
ws = wb.sheets[0]
ws["A1"].options(index=False).value = df

ws["A1"].options(pd.DataFrame, expand="table", index=False).value


# =================================================================
# Pandas Series Converter
# =================================================================

df = pd.read_csv("../summer.csv")
df

athlete = df.Athlete
athlete

wb = xw.Book()
ws = wb.sheets[0]

ws["A1"].value = athlete

ws["B2"].options(pd.Series, expand="table", index=False).value

ws["A1"].expand().clear_contents()

ws["A1"].options(index=False).value = athlete

ws["A1"].options(pd.Series, expand="table", index=False).value

ws["A1"].expand().clear_contents()

df = pd.read_csv("../summer.csv", index_col="Athlete")
df

medal = df.Medal
medal

ws["A1"].value = medal

ws["A1"].options(pd.Series, expand="table", header=False).value
