using XLSX: string
#using Pkg
#Pkg.add("CSV")
#Pkg.add("DataFrames")
#Pkg.add("XLSX")
using XLSX, DataFrames, CSV, Tables;
IO = XLSX.readdata("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\io-table\\IO.xlsx", "io-table-5!A1:DV130");
SAM = zeros(129,129);
#import numerical data into SAM
SAM[2:115, 2:115] = IO[4:117, 3:116];
SAM[2:115, 122:128] = IO[4:117, 118:124];
SAM[116:121, 2:115] = IO[121:126, 3:116];
SAM = Array{Any}(SAM)
#importing column and row titles into SAM
SAM[1, 2:115] .= SAM[2:115, 1] .= string.(lpad.(string.(IO[1, 3:116]), 4, '0'), string(", "), string.(IO[2, 3:116]));
SAM[1, 122:128] .= SAM[122:128, 1] .= string.(string.(IO[3, 118:124]), string(", "), string.(IO[2, 118:124]));
SAM[1, 116:121] .= SAM[116:121, 1] .= string.(string.(IO[121:126, 1]), string(", "), string.(IO[121:126, 2]));
SAM[129, 1] = SAM[1, 129] = string("Total");