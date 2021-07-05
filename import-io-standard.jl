using Base: Int64
using XLSX: length
#using Pkg
#Pkg.add("CSV")
#Pkg.add("DataFrames")
#Pkg.add("XLSX")
using XLSX, DataFrames, Tables;
IOSource = XLSX.readdata("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\SAMBuilding\\IO.xlsx", "io-table-5!A1:DV130");

#indexing vectors for initial data import groups
intermediaryTotalsCol = findall(x -> occursin("T4", x), string.(IOSource[3,:]));
intermediaryTotalsRow = findall(x -> occursin("T1", x), string.(IOSource[:,1]));
finalTotalsCol = findall(x -> occursin("T6", x), string.(IOSource[3,:]));
finalTotalsRow = findall(x -> occursin("Australian Production", x), string.(IOSource[:,2]));
finalDemandCol = findall(x -> occursin('Q', x), string.(IOSource[3,:]));
factorRow = findall(x -> occursin('P', x), string.(IOSource[:,1]));
IOSourceCol = vcat(intermediaryTotalsCol, finalDemandCol, finalTotalsCol);
IOSourceRow = vcat(intermediaryTotalsRow, factorRow, finalTotalsRow);
#initialising IO
IO = zeros(length(IOSourceRow), length(IOSourceCol));
#import numerical data into IO
IO[1:IORowDims, 1:IOColDims] = IOSource[IOSourceRow, IOSourceCol];
#creating vectors of titles for IO
IOCodeRow = IOSource[IOSourceRow, 1];
IOCodeCol = IOSource[3, IOSourceCol];
IONameRow = IOSource[IOSourceRow, 2];
IONameCol = IOSource[2, IOSourceCol];

#code to sum public and private entities into one collumn
investment = findall(x -> occursin("Capital Formation", x), IONameCol);
IO[:, investment[1]]=sum(eachcol(IO[:, investment[1:2]]));
IO = IO[:,Not(investment[2])];
#alter title vectors accordingly (include Q in total investment collumn in IOcode)
IOCodeCol[investment[1]] = "Q3+Q4";
IOCodeCol = IOCodeCol[Not(investment[2])];
IONameCol[investment[1]] = "Private and Public Gross Fixed Capital Formation";
IONameCol = IONameCol[Not(investment[2])];
#creating a dictionary for the index of each collumn and row in IO by IOCode
IOColDict = Dict(IOCodeCol .=> [1:1:8;]);
IORowDict = Dict(IOCodeRow .=> [1:1:8;]);
IOCapForm = findall(x -> occursin("Capital Formation", x), IONameCol);
IOChangeInv = findall(x -> occursin("Changes in Inventories", x), IONameCol);

#importing relevant ASNA data for table 5
ASNAHouseCap = ExcelReaders.readxl("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\SAMBuilding\\ASNAData\\5204039_Household_Capital_Account.xls", "Data1!A1:T71");
ASNANonFinCap = ExcelReaders.readxl("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\SAMBuilding\\ASNAData\\5204018_NonFin_Corp_Capital_Account.xls", "Data1!A1:T71");
ASNAFinCap = ExcelReaders.readxl("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\SAMBuilding\\ASNAData\\5204026_Fin_Corp_Capital_Account.xls", "Data1!A1:S71");
ASNAGovCap = ExcelReaders.readxl("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\SAMBuilding\\ASNAData\\5204032_GenGov_Capital_Account.xls", "Data1!A1:AV71");
ASNAYearRow = findall(x -> occursin("2019", x), string.(ASNAHouseCap[:,1]));

#creating table 5a - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection a is fixed capital expenditure
table5aNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5aNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total indirect taxes", 
"Total fixed capital expenditure"];
table5aRowDict = Dict(table5aNameRow .=> [1:1:length(table5aNameRow);]);
table5aColDict = Dict(table5aNameCol .=> [1:1:length(table5aNameCol);]);
table5a = zeros(length(table5aNameRow), length(table5aNameCol));

#filling in totals collumn from corresponding IO data
table5a[table5aRowDict["Domestic Commodities"], table5aColDict["Total"]] = sum(IO[IORowDict["T1"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, complementary"], table5aColDict["Total"]] = sum(IO[IORowDict["P5"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, competing"], table5aColDict["Total"]] = sum(IO[IORowDict["P6"],IOCapForm]);
table5a[table5aRowDict["Taxes less subsidies on products"], table5aColDict["Total"]] = sum(IO[IORowDict["P3"],IOCapForm]);
table5a[table5aRowDict["Other taxes less subsidies on investment"], table5aColDict["Total"]] = sum(IO[IORowDict["P4"],IOCapForm]);
table5aTaxes = findall(x -> occursin("taxes", x), table5aNameRow);
table5a[table5aRowDict["Total indirect taxes"], table5aColDict["Total"]] = sum(table5a[table5aTaxes,table5aColDict["Total"]]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Total"]] = sum(table5a[Not(table5aRowDict["Total indirect taxes"]),table5aColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapTotCapForm = findall(x -> occursin("General government ;  Gross fixed capital formation ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapTotCapForm]);

#filling in non-total values
for ring in [1:1:length(table5aColDict)-1;];
    table5a[table5aRowDict["Domestic Commodities"],ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*IO[IORowDict["T1"],IOCapForm[1]] / IO[IORowDict["Missing"],IOCapForm[1]];

#creating table 5b - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection b is fixed capital expenditure
table5bNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5bNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Total change in inventories"];
table5bRowDict = Dict(table5bNameRow .=> [1:1:length(table5bNameRow);]);
table5bColDict = Dict(table5bNameCol .=> [1:1:length(table5bNameCol);]);
table5b = zeros(length(table5bNameRow), length(table5bNameCol));

#filling in totals collumn from corresponding IO data
table5b[table5bRowDict["Domestic Commodities"], table5bColDict["Total"]] = sum(IO[IORowDict["T1"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, complementary"], table5bColDict["Total"]] = sum(IO[IORowDict["P5"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, competing"], table5bColDict["Total"]] = sum(IO[IORowDict["P6"],IOChangeInv]);
table5b[table5bRowDict["Taxes less subsidies on products"], table5bColDict["Total"]] = sum(IO[IORowDict["P3"],IOChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Total"]] = sum(table5b[:,table5bColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapChangeInv = findall(x -> occursin("General government ;  Changes in inventories ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapChangeInv]);

#creating table 5c - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection c is totals
table5cNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5cNameRow = ["Domestic Commodities", "Imported Commodities", "Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total investment expenditure"];
table5cRowDict = Dict(table5cNameRow .=> [1:1:length(table5cNameRow);]);
table5cColDict = Dict(table5cNameCol .=> [1:1:length(table5cNameCol);]);
table5c = zeros(length(table5cNameRow), length(table5cNameCol));

#do totals calcuations to get all values

#=convert dataframe to dictionary
function increment!( d::Dict{S, T}, k::S, i::T) where {T<:Real, S<:Any}
    if haskey(d, k)
        d[k] += i
    else
        d[k] = i
    end
end
increment!(d::Dict{S, T}, k::S ) where {T<:Real, S<:Any} = increment!( d, k, one(T))

function df2dict( df::DataFrame, key_col::Symbol, val_col::Symbol=:null)
    keytype = typeof(df[1,key_col])
    if val_col == :null
        valtype = Int
    else
        valtype = typeof(df[1,val_col])
    end
    D = Dict{keytype, valtype}()
    for i=1:size(df,1)
        if !ismissing(df[i,key_col])
            if val_col == :null
                increment!( D, df[i,key_col] )
            elseif valtype <: Real
                increment!( D, df[i,key_col], df[i,val_col] )
            else
                if haskey(D, df[i,key_col])
                    @warn("non-unique entry: $(df[i,key_col])")
                else
                    D[df[i,key_col]] = df[i,val_col]
                end
            end
        end
    end
    return D
end
df[!, "IOCode"]=IOcode
insertcols!(df, 2, :name => vector)
D = df2dict(df, :IOCode, :x3)
=#

