#using Pkg
#Pkg.add("CSV")
#Pkg.add("DataFrames")
#Pkg.add("XLSX")
using XLSX
using DataFrames
IO = XLSX.readdata("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\io-table\\IO.xlsx", "io-table-5!A1:DV130")
df = 