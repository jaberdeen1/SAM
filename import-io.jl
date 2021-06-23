using Base: Int64
using XLSX: length
#using Pkg
#Pkg.add("CSV")
#Pkg.add("DataFrames")
#Pkg.add("XLSX")
using XLSX, DataFrames, Tables, SymPy;
IO = XLSX.readdata("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\io-table\\IO.xlsx", "io-table-5!A1:DV130");

#indexing vectors for initial data import groups
intermediaryCol = findall(x -> occursin("0", x), string.(IO[1,:]));
intermediaryRow = findall(x -> occursin("0", x), string.(IO[:,1]));
finalDemandCol = findall(x -> occursin('Q', x), string.(IO[3,:]));
factorRow = findall(x -> occursin('P', x), string.(IO[:,1]));
#initialising SAM
SAMDims = length(intermediaryCol) + length(finalDemandCol) + length(factorRow) + 1;
SAM = zeros(SAMDims,SAMDims);
#import numerical data into SAM
SAM[1:length(intermediaryRow), 1:length(intermediaryCol)] = IO[intermediaryRow, intermediaryCol];
SAM[1:length(intermediaryRow), length(intermediaryCol)+1:length(intermediaryCol)+ length(finalDemandCol)] = IO[intermediaryRow, finalDemandCol];
SAM[length(intermediaryRow) + 1 : length(intermediaryCol) + length(factorRow), intermediaryCol] = IO[factorRow, intermediaryCol];
#creating vectors of titles for SAM
IOCode = append!(string.(IO[1, intermediaryCol]), string.(IO[factorRow, 1]), string.(IO[3,finalDemandCol]), string.(["Total"]));
IOName = append!(string.(IO[2, intermediaryCol]), string.(IO[factorRow, 2]), string.(IO[2,finalDemandCol]), string.(["Total"]));
#=code to sum all investment final demand into one column, comment out if not needed
#convert to dataframe, make 1st cap formation column into the sum collumn, delete other collumns, delete other rows
investment = findall(x -> occursin("Capital Formation", x), string.(IO[2,:]));
SAM = DataFrame(SAM, :auto);
names!(SAM, Symbol.(IOCode));
SAM[:,investment[1]]=sum(eachcol(SAM[:,investment]));
SAM = SAM[Not(investment[Not(investment[1])]),Not(investment[Not(investment[1])])];
#alter title vectors accordingly (include Q in total investment collumn in IOcode)
IOCode[investment[1]]
IOCode = IOCode[Not(investment[Not(investment[1])])];
IONames = IONames[Not(investment[Not(investment[1])])];
#convert back to matrix
=#
#combining capital formation collumns into relevant collumns
#lambdaPub is the government owned share of public corporations (min 0.5)
#lambdaPriv is the government owned share of private corporations (max 0.5)
@vars lambdaPub
@vars lambdaPriv
#SAM = Any[SAM]
SAM = DataFrame(SAM, :auto);
privInv = findall(x -> occursin("Private", x), IOName);
pubInv = findall(x -> occursin("Public", x), IOName)[3];
govInv = findall(x -> occursin("Government", x), IOName)[2];
householdsExpend = findall(x -> occursin("Households", x), IOName);
govExpend = findall(x -> occursin("Government", x), IOName)[1];
for i in 1:length(IOName);
    SAM[i, govExpend]=SAM[i, govExpend]+lambdaPub*SAM[i,pubInv]+lambdaPriv*SAM[i,privInv]+SAM[i, govInv]; 
    SAM[i, householdsExpend]=SAM[i, householdsExpend]+(1-lambdaPub)*pubInv+(1-lambdaPriv)*privInv; 
end

#SAM = SAM[Not(investment[Not(investment[1])]),Not(investment[Not(investment[1])])];
#alter title vectors accordingly (include Q in total investment collumn in IOcode)
IOCode[investment[1]]
IOCode = IOCode[Not(investment[Not(investment[1])])];
IONames = IONames[Not(investment[Not(investment[1])])];
#creating vectors for groups within SAM
intermediary = findall(x -> occursin("0", x), IOcode);
finalDemand = findall(x -> occursin("Q", x), IOcode);
factor = findall(x -> occursin("P", x), IOcode);
#steps to compute missing values
#for ring in 1:127;
 #   SAM[ring,128] = sum(SAM[ring,1:127]);
  #  SAM[128,ring] = sum(SAM[1:127,ring]);
#end
#assuming compensation and gross operating surplus all goes into households;
#SAM[122, 116:117] .= SAM[116:117, 129];
#assuming all tax goes directly to the government;
#SAM[12];

#=
#convert to dictionary
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
D = df2dict(df, :IOCode, :x3)
=#
