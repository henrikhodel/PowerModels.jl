# import packages/
using CSV, DataFrames, XLSX
using JSONTables, JSON, DataStructures, Plots
import DataStructures: OrderedDict

hours = 1:8760

# import buses and lines as dataframe in julia
global df_bus = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/buses_post_joint_no_dc.csv", DataFrame)
df_lines = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/lines_post_joint.csv", DataFrame)
df_links = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/links.csv", DataFrame)
df_trafo = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/transformers.csv", DataFrame)
df_branch = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/lines_prep_for_PM.xlsx","to_julia"))

# import load and generation factors from xlsx as dataframe in julia
df_load_factors_SE = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/SE_load.xlsx","to_julia"))
df_load_factors_DK = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/DK_load.xlsx","to_julia"))
df_load_factors_FI = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/FI_load.xlsx","to_julia"))
df_load_factors_NO = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/NO_load.xlsx","to_julia"))

df_gen_factors_SEs = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/SE_gen.xlsx","to_julia"))
df_gen_factors_DKs = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/DK_gen.xlsx","to_julia"))
df_gen_factors_FIs = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/FI_gen.xlsx","to_julia"))
df_gen_factors_NOs = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/NO_gen.xlsx","to_julia"))

# import loads and generation timeseries as dataframe in julia
global df_loads = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/Processed data/Load_nordic_adjusted.csv", DataFrame)

df_load_SE1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_1_load.csv", DataFrame)
df_load_SE2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_2_load.csv", DataFrame)
df_load_SE3 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_3_load.csv", DataFrame)
df_load_SE4 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_4_load.csv", DataFrame)

df_load_NO1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_1_load.csv", DataFrame)
df_load_NO2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_2_load.csv", DataFrame)
df_load_NO3 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_3_load.csv", DataFrame)
df_load_NO4 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_4_load.csv", DataFrame)
df_load_NO5 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_5_load.csv", DataFrame)

df_load_DK1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/DK_1_load.csv", DataFrame)
df_load_DK2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/DK_2_load.csv", DataFrame)

df_load_FI = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/FI_load.csv", DataFrame)

# import import and export timeseries as dataframe in julia
df_crossborder = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/nordic_dc_lines_import_export.xlsx","master"))

# generation
df_gen_SE1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_1_generation.csv", DataFrame)
df_gen_SE2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_2_generation.csv", DataFrame)
df_gen_SE3 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_3_generation.csv", DataFrame)
df_gen_SE4 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/SE_4_generation.csv", DataFrame)

df_gen_NO1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_1_generation.csv", DataFrame)
df_gen_NO2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_2_generation.csv", DataFrame)
df_gen_NO3 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_3_generation.csv", DataFrame)
df_gen_NO4 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_4_generation.csv", DataFrame)
df_gen_NO5 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/NO_5_generation.csv", DataFrame)

df_gen_DK1 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/DK_1_generation.csv", DataFrame)
df_gen_DK2 = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/DK_2_generation.csv", DataFrame)

df_gen_FI = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/FI_generation.csv", DataFrame)

df_gen_factors_other_DK = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/Other_gen.xlsx","to_julia_DK"))
df_gen_factors_other_FI = DataFrame(XLSX.readtable("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/load and gen allocation/Other_gen.xlsx","to_julia_FI"))
display("successfully loaded data")

df_load_SE1 = coalesce.(df_load_SE1, 0)
df_load_SE2 = coalesce.(df_load_SE2, 0)
df_load_SE3 = coalesce.(df_load_SE3, 0)
df_load_SE4 = coalesce.(df_load_SE4, 0)

df_load_NO1 = coalesce.(df_load_NO1, 0)
df_load_NO2 = coalesce.(df_load_NO2, 0)
df_load_NO3 = coalesce.(df_load_NO3, 0)
df_load_NO4 = coalesce.(df_load_NO4, 0)
df_load_NO5 = coalesce.(df_load_NO5, 0)

df_load_DK1 = coalesce.(df_load_DK1, 0)
df_load_DK2 = coalesce.(df_load_DK2, 0)

df_load_FI = coalesce.(df_load_FI, 0)

df_gen_SE1 = coalesce.(df_gen_SE1, 0)
df_gen_SE2 = coalesce.(df_gen_SE2, 0)
df_gen_SE3 = coalesce.(df_gen_SE3, 0)
df_gen_SE4 = coalesce.(df_gen_SE4, 0)

df_gen_NO1 = coalesce.(df_gen_NO1, 0)
df_gen_NO2 = coalesce.(df_gen_NO2, 0)
df_gen_NO3 = coalesce.(df_gen_NO3, 0)
df_gen_NO4 = coalesce.(df_gen_NO4, 0)
df_gen_NO5 = coalesce.(df_gen_NO5, 0)

df_gen_DK1 = coalesce.(df_gen_DK1, 0)
df_gen_DK2 = coalesce.(df_gen_DK2, 0)

df_gen_FI = coalesce.(df_gen_FI, 0)

df_gen_factors_other_DK = coalesce.(df_gen_factors_other_DK, 0)
df_gen_factors_other_FI = coalesce.(df_gen_factors_other_FI, 0)

# fix some incorrect generation

df_gen_SE2[2873:2879,:"Hydro Water Reservoir"] .= [3662, 3779, 3959, 3842, 3834, 3489, 3405]
df_gen_SE2[8131:8136,:"Hydro Water Reservoir"] .= [5407, 5297, 5225, 5034, 4922, 4633]


# prepare the bus dataframe
df_bus = select(df_bus,["name", "v_nom"])
df_bus = rename(df_bus, [1 => :"bus_i", 2 => :"base_kv"])
df_bus[!,"bus_type"] = ones(size(df_bus,1))
df_bus[!,"vmin"] = ones(size(df_bus,1)) .* 0.95 
df_bus[!,"vmax"] = ones(size(df_bus,1)) .* 1.05
df_bus[!,"va"] = ones(size(df_bus,1)) .* 0
df_bus[!,"vm"] = ones(size(df_bus,1))

# buses that belong to DK1
# manually
buses_DK1_manually = [
    5802,
    5803,
    5805,
    5815,
    5818,
    5821,
    5822,
    5824,
    5836,
    5838,
    5842,
    5847,
    6369,
    6385,
    6388,
    6398,
    6402,
    6408,
    6410,
    6416,
    6421,
    6424,
    6425,
    6431,
    6434,
    7332,
    7333,
    7365,
    7905,
    8567
]


buses_DK1_load = df_load_factors_DK[df_load_factors_DK.load_factor_DK1 .> 0, :name]
buses_DK1_gen = df_gen_factors_DKs[df_gen_factors_DKs.Wind_DK1 .> 0, :name]

buses_DK1_auto = union(buses_DK1_load, buses_DK1_gen)
buses_DK1_auto = parse.(Int,string.(buses_DK1_auto))
buses_DK1_all = union(buses_DK1_auto, buses_DK1_manually)
# remove rows of buses that do not belong in the model of the synchronous Grid (DK1 + other ends of dc links)
outside_buses = [
    6485,
    5949,
    7909,
    5749,
    5751,
    5858,
    5685,
    5686,
    5588,
    5605,
    6558,
    6577
]
other_buses = [
    6300,
    7430
]
# Gotland (6300)

# joint for Kriegers flak (7430)

# nordic buses = all buses - DK1_manual buses - outside buses - others buses to be removed
all_buses = df_bus[:,1]
buses_nordic = symdiff(all_buses, buses_DK1_all)
symdiff!(buses_nordic, outside_buses)

df_load_factors_SE[!,:name] = parse.(Int,string.(df_load_factors_SE[!,:name]));
df_load_factors_NO[!,:name] = parse.(Int,string.(df_load_factors_NO[!,:name]));
df_load_factors_DK[!,:name] = parse.(Int,string.(df_load_factors_DK[!,:name]));
df_load_factors_FI[!,:name] = parse.(Int,string.(df_load_factors_FI[!,:name]));

df_gen_factors_SEs[!,:name] = parse.(Int,string.(df_gen_factors_SEs[!,:name]));
df_gen_factors_NOs[!,:name] = parse.(Int,string.(df_gen_factors_NOs[!,:name]));
df_gen_factors_DKs[!,:name] = parse.(Int,string.(df_gen_factors_DKs[!,:name]));
df_gen_factors_FIs[!,:name] = parse.(Int,string.(df_gen_factors_FIs[!,:name]));

df_gen_factors_other_DK[!,:name] = parse.(Int,string.(df_gen_factors_other_DK[!,:name]));
df_gen_factors_other_FI[!,:name] = parse.(Int,string.(df_gen_factors_other_FI[!,:name]));


# allocate the loads (need to remove the losses as they will be calculated in the model)
# mulitply the load factor value with the actual load value for rows with the same bus in the load factor and load dataframes
# then sum the load values for each bidding zone

# for every hour in the df_load dataframes, multiply the load factor with the load value for each bidding zone

# vector of hours in the year

# create a dataframe and set bus from df_buses as row index

df_pd = DataFrame(bus = df_bus[:,1])
df_pd_SE = DataFrame(bus = df_load_factors_SE[:,1])
df_pd_NO = DataFrame(bus = df_load_factors_NO[:,1])
df_pd_DK = DataFrame(bus = df_load_factors_DK[:,1])
df_pd_FI = DataFrame(bus = df_load_factors_FI[:,1])

# add column called "qd" to the dataframes and fill it with zeroes
df_pd_SE.pd = zeros(size(df_pd_SE,1))
df_pd_NO.pd = zeros(size(df_pd_NO,1))
df_pd_DK.pd = zeros(size(df_pd_DK,1))
df_pd_FI.pd = zeros(size(df_pd_FI,1))

# for every bus in df_qd, sum the product of the load factor and the load value for each bidding zone
# then add the sum to the dataframe df_pd
# make a vector containing the buses of the column 1 of df_pd
buses = df_pd[:,1]
#buses = string.(buses)
#buses = Vector{Any}(buses)

buses_gen_SE = union(df_gen_factors_SEs[:,1],[6348, 6312, 6591])
buses_gen_NO = df_gen_factors_NOs[:,1]
buses_gen_DK = union(df_gen_factors_DKs[:,1],df_gen_factors_other_DK[:,1])
buses_gen_FI = union(df_gen_factors_FIs[:,1],df_gen_factors_other_FI[:,1], [6617, 6570])

buses_load_DK = df_load_factors_DK[:,1]
buses_load_FI = df_load_factors_FI[:,1]
buses_load_NO = df_load_factors_NO[:,1]
buses_load_SE = df_load_factors_SE[:,1]

SE1_load_loss_factor = 0 # 2.2 # % loss 
SE2_load_loss_factor = 0 # 11.4 # 16.1 # % loss
SE3_load_loss_factor = 0 # 1 # 4.2 # % loss
SE4_load_loss_factor = 0 # 0.9 # 3 # % loss

NO1_load_loss_factor = 0 # 0 # 2.1 # % loss
NO2_load_loss_factor = 0 # 4.4 # 2.7 # % loss
NO3_load_loss_factor = 0 # 2 # 2.2 # % loss
NO4_load_loss_factor = 0 # 2.5 # 2.8 # % loss
NO5_load_loss_factor = 0 # 4.1 # 4.8 # % loss

DK1_load_loss_factor = 0 # 0 # 1 # % loss
DK2_load_loss_factor = 0 # 0.2 # 0.6 # % loss

FI_load_loss_factor = 0 # 0.6 # 2.6 # % loss

# # # # Generation # # # #

# add the load factors to the gen factors to account for the other and solar generation
df_gen_factors_SE = outerjoin(df_gen_factors_SEs, df_load_factors_SE, on = :name, makeunique = true)
df_gen_factors_NO = outerjoin(df_gen_factors_NOs, df_load_factors_NO, on = :name, makeunique = true)
df_gen_factors_DK = outerjoin(df_gen_factors_DKs, df_load_factors_DK, on = :name, makeunique = true)
df_gen_factors_FI = outerjoin(df_gen_factors_FIs, df_load_factors_FI, on = :name, makeunique = true)
# fill missing fields in df_load_factors with 0
df_gen_factors_SE = coalesce.(df_gen_factors_SE, 0)
df_gen_factors_NO = coalesce.(df_gen_factors_NO, 0)
df_gen_factors_DK = coalesce.(df_gen_factors_DK, 0)
df_gen_factors_FI = coalesce.(df_gen_factors_FI, 0)

df_pg = DataFrame(bus = df_bus[:,1])
df_pg_SE = DataFrame(bus = df_gen_factors_SE[:,1])
df_pg_NO = DataFrame(bus = df_gen_factors_NO[:,1])
df_pg_DK = DataFrame(bus = df_gen_factors_DK[:,1])
df_pg_FI = DataFrame(bus = df_gen_factors_FI[:,1])

df_pg_other_DK = DataFrame(bus = df_gen_factors_other_DK[:,1])
df_pg_other_FI = DataFrame(bus = df_gen_factors_other_FI[:,1])

# add column called "qd" to the dataframes and fill it with zeroes
df_pg_SE.pg = zeros(size(df_pg_SE,1))
df_pg_NO.pg = zeros(size(df_pg_NO,1))
df_pg_DK.pg = zeros(size(df_pg_DK,1))
df_pg_FI.pg = zeros(size(df_pg_FI,1))

df_pg_other_DK.pg_other_DK = zeros(size(df_pg_other_DK,1))
df_pg_other_FI.pg_other_FI = zeros(size(df_pg_other_FI,1))

df_branch[!,:name_str] = string.(df_branch[:,:name])
df_bus[!,:bus_i_str] = string.(df_bus[:,:bus_i])
df_links[!,:name_str] = string.(df_links[:,:name])
df_branch[!,:name] = parse.(Int,string.(df_branch[!,:name]));
df_branch[!,:f_bus] = parse.(Int,string.(df_branch[!,:f_bus]));
df_branch[!,:t_bus] = parse.(Int,string.(df_branch[!,:t_bus]));
df_links[!,:name] = parse.(Int,string.(df_links[!,:name]));

df_bus_nordic = df_bus[in.(df_bus.bus_i, Ref(buses_nordic)), :] # only nordic buses
df_branch_nordic = df_branch[in.(df_branch.f_bus, Ref(buses_nordic)), :] # only nordic buses

# 7003 = hårsprånget, 6591 = forsmark
df_bus_nordic[df_bus_nordic.bus_i.==6591,:bus_type] .= 3 # set the reference bus for the nordics

df_bus_DK1 = df_bus[in.(df_bus.bus_i, Ref(buses_DK1_all)), :] # only DK1
df_branch_DK1 = df_branch[in.(df_branch.f_bus, Ref(buses_DK1_all)), :] # only DK1

df_bus_DK1[df_bus_DK1.bus_i.==6402,:bus_type] .= 3 # set the reference bus for DK1


# start imbalance loop here
imbalance = []
gens = []
net_exps = []
loads = []
for hour in hours
    for bus in buses
        if bus in df_load_factors_SE[:,1]
            local df_pd_SE[df_pd_SE.bus.==bus,"pd"] = (df_load_factors_SE[df_load_factors_SE.name.==bus,"load_factor_SE1"] * df_loads[hour,"SE1"] * (1-SE1_load_loss_factor/100)
                + df_load_factors_SE[df_load_factors_SE.name.==bus,"load_factor_SE2"] * df_loads[hour,"SE2"] * (1-SE2_load_loss_factor/100)
                + df_load_factors_SE[df_load_factors_SE.name.==bus,"load_factor_SE3"] * df_loads[hour,"SE3"] * (1-SE3_load_loss_factor/100)
                + df_load_factors_SE[df_load_factors_SE.name.==bus,"load_factor_SE4"] * df_loads[hour,"SE4"] * (1-SE4_load_loss_factor/100)
            )

        elseif bus in df_load_factors_NO[:,1]
            local df_pd_NO[df_pd_NO.bus.==bus,"pd"] = (df_load_factors_NO[df_load_factors_NO.name.==bus,"load_factor_NO1"] * df_loads[hour,"NO1"] * (1-NO1_load_loss_factor/100)
                + df_load_factors_NO[df_load_factors_NO.name.==bus,"load_factor_NO2"] * df_loads[hour,"NO2"] * (1-NO2_load_loss_factor/100)
                + df_load_factors_NO[df_load_factors_NO.name.==bus,"load_factor_NO3"] * df_loads[hour,"NO3"] * (1-NO3_load_loss_factor/100)
                + df_load_factors_NO[df_load_factors_NO.name.==bus,"load_factor_NO4"] * df_loads[hour,"NO4"] * (1-NO4_load_loss_factor/100)
                + df_load_factors_NO[df_load_factors_NO.name.==bus,"load_factor_NO5"] * df_loads[hour,"NO5"] * (1-NO5_load_loss_factor/100)
            )

        elseif bus in df_load_factors_DK[:,1]
            local df_pd_DK[df_pd_DK.bus.==bus,"pd"] = (df_load_factors_DK[df_load_factors_DK.name.==bus,"load_factor_DK1"] * df_loads[hour,"DK1"] * (1-DK1_load_loss_factor/100)
                + df_load_factors_DK[df_load_factors_DK.name.==bus,"load_factor_DK2"] * df_loads[hour,"DK2"] * (1-DK2_load_loss_factor/100)
            )
    
        elseif bus in df_load_factors_FI[:,1]
            local df_pd_FI[df_pd_FI.bus.==bus,"pd"] = df_load_factors_FI[df_load_factors_FI.name.==bus,"load_factor_FI"] * df_loads[hour,"FI"] * (1-FI_load_loss_factor/100)
        else
        end
    end
    # 

    # merge load factor dataframes
    local df_load = outerjoin(df_pd_SE, df_pd_DK, df_pd_FI, df_pd_NO, on = :bus, makeunique = true)
    # fill missing fields in df_load_factors with 0
    local df_load = coalesce.(df_load, 0)
    # sum the pd columns
    local df_load[!,"pdsum"] = df_load[!,"pd"] + df_load[!,"pd_1"] + df_load[!,"pd_2"] + df_load[!,"pd_3"]
    # keep only the bus and pdsum columns
    local df_load = select(df_load,["bus", "pdsum"])
    # rename the pdsum column to pd
    local df_load = rename(df_load, [1 => :"load_bus", 2 => :"pd"])

    # set the load power factor
    #reactive_power_factor_load = 0

    # calculate the reactive power demand
    #df_load[!,"qd"] = df_load[!,"pd"] .* reactive_power_factor_load

    local df_gens_SE = sum(df_gen_SE1[hour,2:end]) + sum(df_gen_SE2[hour,2:end]) + sum(df_gen_SE3[hour,2:end]) + sum(df_gen_SE4[hour,2:end])
    local df_gens_NO = sum(df_gen_NO1[hour,2:end]) + sum(df_gen_NO2[hour,2:end]) + sum(df_gen_NO3[hour,2:end]) + sum(df_gen_NO4[hour,2:end]) + sum(df_gen_NO5[hour,2:end])
    local df_gens_DK = sum(df_gen_DK1[hour,2:end]) + sum(df_gen_DK2[hour,2:end])
    local df_gens_FI = sum(df_gen_FI[hour,2:end])

    for bus in buses
        # println("bus $bus")
        if bus in df_gen_factors_SE[:,1]
            local df_pg_SE[df_pg_SE.bus.==bus,"pg"] =
            (
                df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Wind_SE1"] * df_gen_SE1[hour,"Wind Onshore"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Wind_SE2"] * df_gen_SE2[hour,"Wind Onshore"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Wind_SE3"] * df_gen_SE3[hour,"Wind Onshore"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Wind_SE4"] * df_gen_SE4[hour,"Wind Onshore"]
            
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Hydro_SE1"] * df_gen_SE1[hour,"Hydro Water Reservoir"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Hydro_SE2"] * df_gen_SE2[hour,"Hydro Water Reservoir"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Hydro_SE3"] * df_gen_SE3[hour,"Hydro Water Reservoir"]
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"Hydro_SE4"] * df_gen_SE4[hour,"Hydro Water Reservoir"]

                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"load_factor_SE1"] * (df_gen_SE1[hour,"Other"] + df_gen_SE1[hour,"Solar"])
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"load_factor_SE2"] * (df_gen_SE2[hour,"Other"] + df_gen_SE2[hour,"Solar"] + df_gen_SE2[hour,"Fossil Gas"])
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"load_factor_SE3"] * (df_gen_SE3[hour,"Other"] + df_gen_SE3[hour,"Solar"] + df_gen_SE3[hour,"Fossil Gas"])
                + df_gen_factors_SE[df_gen_factors_SE.name.==bus,"load_factor_SE4"] * (df_gen_SE4[hour,"Other"] + df_gen_SE4[hour,"Solar"] + df_gen_SE4[hour,"Fossil Gas"])
            )

        elseif bus in df_gen_factors_NO[:,1]
            local df_pg_NO[df_pg_NO.bus.==bus,"pg"] =
            (
                df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Wind_NO1"] * df_gen_NO1[hour,"Wind Onshore"]
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Wind_NO2"] * df_gen_NO2[hour,"Wind Onshore"]
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Wind_NO3"] * df_gen_NO3[hour,"Wind Onshore"]
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Wind_NO4"] * df_gen_NO4[hour,"Wind Onshore"]
            
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Hydro_NO1"] * (df_gen_NO1[hour,"Hydro Water Reservoir"] + df_gen_NO1[hour,"Hydro Run-of-river and poundage"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Hydro_NO2"] * (df_gen_NO2[hour,"Hydro Water Reservoir"] + df_gen_NO2[hour,"Hydro Run-of-river and poundage"] + df_gen_NO2[hour,"Hydro Pumped Storage"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Hydro_NO3"] * (df_gen_NO3[hour,"Hydro Water Reservoir"] + df_gen_NO3[hour,"Hydro Run-of-river and poundage"] + df_gen_NO3[hour,"Hydro Pumped Storage"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Hydro_NO4"] * (df_gen_NO4[hour,"Hydro Water Reservoir"] + df_gen_NO4[hour,"Hydro Run-of-river and poundage"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"Hydro_NO5"] * (df_gen_NO5[hour,"Hydro Water Reservoir"] + df_gen_NO5[hour,"Hydro Run-of-river and poundage"] + df_gen_NO5[hour,"Hydro Pumped Storage"])

                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"load_factor_NO1"] * (df_gen_NO1[hour,"Biomass"] + df_gen_NO1[hour,"Waste"] + df_gen_NO1[hour,"Fossil Gas"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"load_factor_NO2"] * (df_gen_NO2[hour,"Waste"] + df_gen_NO2[hour,"Fossil Gas"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"load_factor_NO3"] * (df_gen_NO3[hour,"Waste"] + df_gen_NO3[hour,"Fossil Gas"] + df_gen_NO3[hour,"Other renewable"] + df_gen_NO3[hour,"Other"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"load_factor_NO4"] * (df_gen_NO4[hour,"Other renewable"] + df_gen_NO4[hour,"Fossil Gas"])
                + df_gen_factors_NO[df_gen_factors_NO.name.==bus,"load_factor_NO5"] * (df_gen_NO5[hour,"Waste"] + df_gen_NO5[hour,"Fossil Gas"])
            )

        elseif bus in df_gen_factors_DK[:,1]
            local df_pg_DK[df_pg_DK.bus.==bus,"pg"] =
            (
                df_gen_factors_DK[df_gen_factors_DK.name.==bus,"Wind_DK1"] * df_gen_DK1[hour,"Wind Onshore"]
                + df_gen_factors_DK[df_gen_factors_DK.name.==bus,"Wind_DK2"] * df_gen_DK2[hour,"Wind Onshore"]
            
                + df_gen_factors_DK[df_gen_factors_DK.name.==bus,"load_factor_DK1"] * (df_gen_DK1[hour,"Other renewable"] + df_gen_DK1[hour,"Solar"] + df_gen_DK1[hour,"Hydro Run-of-river and poundage"] + df_gen_DK1[hour,"Fossil Oil"] + df_gen_DK1[hour,"Waste"])
                + df_gen_factors_DK[df_gen_factors_DK.name.==bus,"load_factor_DK2"] * (df_gen_DK2[hour,"Solar"] + df_gen_DK2[hour,"Waste"] + df_gen_DK2[hour,"Fossil Hard coal"] + df_gen_DK2[hour,"Fossil Gas"])
            )
            # Biomass, Fossil Gas, Fossil Hard coal and Wind Offshore are manually assigned to buses
        elseif bus in df_gen_factors_FI[:,1]
            local df_pg_FI[df_pg_FI.bus.==bus,"pg"] =
            (
                df_gen_factors_FI[df_gen_factors_FI.name.==bus,"wind_factor"] * df_gen_FI[hour,"Wind Onshore"]

                + df_gen_factors_FI[df_gen_factors_FI.name.==bus,"hydro_factor"] * df_gen_FI[hour,"Hydro Run-of-river and poundage"]
            
                + df_gen_factors_FI[df_gen_factors_FI.name.==bus,"load_factor_FI"] * (df_gen_FI[hour,"Other renewable"] + df_gen_FI[hour,"Solar"] + df_gen_FI[hour,"Waste"])
            )
            # Biomass, Fossil Gas, Fossil Hard coal, Fossil Peat and Nuclear are manually assigned to buses
        else
        end
    end

    for bus in buses
        # println("bus $bus")
        if bus in df_gen_factors_other_DK[:,1]
            local df_pg_other_DK[df_pg_other_DK.bus.==bus,"pg_other_DK"] =
            (
                df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Wind Offshore_DK1"] * df_gen_DK1[hour,"Wind Offshore"]
                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Wind Offshore_DK2"] * df_gen_DK2[hour,"Wind Offshore"]

                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Fossil Hard coal_DK1"] * df_gen_DK1[hour,"Fossil Hard coal"]

                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Fossil Gas_DK1"] * df_gen_DK1[hour,"Fossil Gas"]

                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Fossil Oil_DK2"] * df_gen_DK2[hour,"Fossil Oil"]

                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Biomass_DK1"] * df_gen_DK1[hour,"Biomass"]
                + df_gen_factors_other_DK[df_gen_factors_other_DK.name.==bus,"Biomass_DK2"] * df_gen_DK2[hour,"Biomass"]
            )
        elseif bus in df_gen_factors_other_FI[:,1]
            local df_pg_other_FI[df_pg_other_FI.bus.==bus,"pg_other_FI"] =
            (
                df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Fossil Peat"] * df_gen_FI[hour,"Fossil Peat"]
                + df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Fossil Hard coal"] * df_gen_FI[hour,"Fossil Hard coal"]
                + df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Fossil Gas"] * df_gen_FI[hour,"Fossil Gas"]
                + df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Fossil Oil"] * df_gen_FI[hour,"Fossil Oil"]
                + df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Biomass"] * df_gen_FI[hour,"Biomass"]
                + df_gen_factors_other_FI[df_gen_factors_other_FI.name.==bus,"Other"] * df_gen_FI[hour,"Other"]
            )
        else
        end
    end

    # merge load factor dataframes
    local df_gen = DataFrame(gen_bus = Int64[], pg = Float64[])
    local df_gen = outerjoin(df_pg_SE, df_pg_DK, df_pg_FI, df_pg_NO, df_pg_other_DK, df_pg_other_FI, on = :bus, makeunique = true)
    # fill missing fields in df_load_factors with 0
    local df_gen = coalesce.(df_gen, 0)
    # sum the pd columns
    local df_gen[!,"pgsum"] = df_gen[!,"pg"] + df_gen[!,"pg_1"] + df_gen[!,"pg_2"] + df_gen[!,"pg_3"] + df_gen[!,"pg_other_DK"] + df_gen[!,"pg_other_FI"]
    # keep only the bus and pdsum columns
    local df_gen = select(df_gen,["bus", "pgsum"])
    # rename the pdsum column to pd
    local df_gen = rename(df_gen, [1 => :"gen_bus", 2 => :"pg"])


    # hardcode nuclear and thermal generation to the right buses 
    # quick and dirty for now

    # nuclear generators
    local SE_nuc_tot = 1380+2193+3283
    # orskarshamn 3 (1 380 MW) = 6348
    local df_gen = vcat(df_gen, DataFrame(gen_bus = 6348, pg = ((1380/SE_nuc_tot) * df_gen_SE3[hour,"Nuclear"])))
    # ringhals 3+4? (2 193MW) = 6312
    local df_gen = vcat(df_gen, DataFrame(gen_bus = 6312, pg = ((2193/SE_nuc_tot) * df_gen_SE3[hour,"Nuclear"])))
    # forsmark 1,2,3 (3 283MW)= 6591
    local df_gen = vcat(df_gen, DataFrame(gen_bus = 6591, pg = ((3283/SE_nuc_tot) * df_gen_SE3[hour,"Nuclear"])))

    local FI_nuc_tot = 3390+1025
    # Olkiluoto = 6617
    local df_gen = vcat(df_gen, DataFrame(gen_bus = 6617, pg = ((3390/FI_nuc_tot) * df_gen_FI[hour,"Nuclear"])))
    # Loviisa = 6570
    local df_gen = vcat(df_gen, DataFrame(gen_bus = 6570, pg = ((1025/FI_nuc_tot) * df_gen_FI[hour,"Nuclear"])))

    local df_load[!,:load_bus] = parse.(Int,string.(df_load[!,:load_bus]));

    local df_load = combine(groupby(df_load, :load_bus), :pd => sum, renamecols=false)
    local df_gen = combine(groupby(df_gen, :gen_bus), :pg => sum, renamecols=false)


    # allocate the imports and exports as load to the corresponding buses
    #println("allocating the import/export to the corresponding buses: ")
    #println("load before:", sum(df_load[!,"pd"]))
    #println("gen:", sum(df_gen[!, :pg]))

    local df_net_exp = DataFrame(load_bus = Int64[], pd = Float64[])
    # DK1 - NO2
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6398, pd = float(df_crossborder[hour,"DK1 > NO2"] - df_crossborder[hour,"DK1 < NO2"]))) # DK side
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6365, pd = float(df_crossborder[hour,"NO2 > DK1"] - df_crossborder[hour,"NO2 < DK1"]))) # NO side

    # DK1 - DE
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5822, pd = float(df_crossborder[hour,"DK1 > DE"] - df_crossborder[hour,"DK1 < DE"]))) # DK side

    # DK1 - SE3
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6416, pd = float(df_crossborder[hour,"DK1 > SE3"] - df_crossborder[hour,"DK1 < SE3"]))) # DK side
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6311, pd = float(df_crossborder[hour,"SE3 > DK1"] - df_crossborder[hour,"SE3 < DK1"]))) # SE side

    # DK1 - NL
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5805, pd = float(df_crossborder[hour,"DK1 > NL"] - df_crossborder[hour,"DK1 < NL"]))) # DK side

    # DK1 - DK2
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5847, pd = float(df_crossborder[hour,"DK1 > DK2"] - df_crossborder[hour,"DK1 < DK2"]))) # DK1 side
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5624, pd = float(df_crossborder[hour,"DK2 > DK1"] - df_crossborder[hour,"DK2 < DK1"]))) # DK2 side

    # DK2 - DE
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5646, pd = float(df_crossborder[hour,"DK2 > DE"] - df_crossborder[hour,"DK2 < DE"]))) # DK2 side

    # NO2 - DE
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6358, pd = float(df_crossborder[hour,"NO2 > DE"] - df_crossborder[hour,"NO2 < DE"]))) # NO2 side

    # NO2 - GB
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6695, pd = float(df_crossborder[hour,"NO2 > GB"] - df_crossborder[hour,"NO2 < GB"]))) # NO2 side

    # NO2 - NL
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6363, pd = float(df_crossborder[hour,"NO2 > NL"] - df_crossborder[hour,"NO2 < NL"]))) # NO2 side

    # SE3 - FI
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6598, pd = float(df_crossborder[hour,"SE3 > FI"] - df_crossborder[hour,"SE3 < FI"])/2)) # SE3 side (half of the import)
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6586, pd = float(df_crossborder[hour,"SE3 > FI"] - df_crossborder[hour,"SE3 < FI"])/2)) # SE3 side (other half of the import)
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6618, pd = float(df_crossborder[hour,"FI > SE3"] - df_crossborder[hour,"FI < SE3"]))) # FI side

    # SE4 - DE
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 5673, pd = float(df_crossborder[hour,"SE4 > DE"] - df_crossborder[hour,"SE4 < DE"]))) # SE4 side

    # SE4 - LT
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6354, pd = float(df_crossborder[hour,"SE4 > LT"] - df_crossborder[hour,"SE4 < LT"]))) # SE4 side

    # SE4 - PL
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6351, pd = float(df_crossborder[hour,"SE4 > PL"] - df_crossborder[hour,"SE4 < PL"]))) # SE4 side

    # FI - EE
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6556, pd = float(df_crossborder[hour,"FI > EE"] - df_crossborder[hour,"FI < EE"])/2)) # FI side (half of the import)
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6567, pd = float(df_crossborder[hour,"FI > EE"] - df_crossborder[hour,"FI < EE"])/2)) # FI side (other half of the import)

    # FI - RU
    df_net_exp = vcat(df_net_exp, DataFrame(load_bus = 6573, pd = float(df_crossborder[hour,"FI > RU"] - df_crossborder[hour,"FI < RU"]))) # FI side


    #println("net export:", sum(df_net_exp[!,"pd"]))
    local df_load2 = vcat(df_load, df_net_exp)
    local df_load2 = combine(groupby(df_load2, :load_bus), :pd => sum, renamecols=false)
    #println("load after:", sum(df_load[!,"pd"]))

    # sanity checks
    # check that the generation is equal to the load (including losses)
    local diff = sum(df_load2[!,"pd"]) - sum(df_gen[!,"pg"])
    # display the missing generation
    #print("missing generation: ", diff)
    # cover it by imports (temporary measure)

    # push to here in the for loop

    push!(imbalance, diff)
    push!(gens, round(df_gens_NO+df_gens_SE+df_gens_DK+df_gens_FI-sum(df_gen[!,"pg"]);digits =2))
    push!(loads, round(sum(df_loads[hour,2:13])-sum(df_load[!,"pd"]); digits = 2))

    df_gen[!,:gen_bus_str] = string.(df_gen[:,:gen_bus])
    df_load2[!,:load_bus_str] = string.(df_load2[:,:load_bus])
    df_gen[!,:gen_bus] = parse.(Int,string.(df_gen[!,:gen_bus]));
    
    grid_loss_factor = 1
    # Using JSON.jl to make nested objects
    local load = []
    local df_load_nordic = df_load2[in.(df_load2.load_bus, Ref(buses_nordic)), :] # only nordic buses
    for g in groupby(df_load_nordic, :load_bus)
        load_data = OrderedDict(
            "load_bus" => g.load_bus[end],
            "pd" => g.pd[end] * grid_loss_factor,
            "qd" => 0,
            "index" => g.load_bus[end],
            "status" => 1
        )
        push!(load, [g.load_bus_str[end] load_data])
    end

    local gen = []
    local df_gen_nordic = df_gen[in.(df_gen.gen_bus, Ref(buses_nordic)), :] # only nordic buses
    for g in groupby(df_gen_nordic, :gen_bus)
        gen_data = OrderedDict(
            "gen_bus" => g.gen_bus[end],
            "pg" => g.pg[end],
            "qg" => 0,
            "pmax" => 9999,
            "pmin" => 0,
            "qmax" => 9999,
            "qmin" => -9999,
            "index" => g.gen_bus[end],
            "vg" => 1,
            "gen_status" => 1
        )
        push!(gen, [g.gen_bus_str[end] gen_data])
    end

    bus = []
    for g in groupby(df_bus_nordic, :bus_i)
        bus_data = OrderedDict(
            "bus_type" => g.bus_type[end], # 1 - PQ, 2 - PV, 3 - reference , 4 - isolated  , see https://matpower.org/docs/ref/matpower5.0/caseformat.html
            "vm" => 1,
            "va" => 0,
            "vmin" => 0.5,
            "vmax" => 1.5,
            "index" => g.bus_i[end],
            "gen_status" => 1,
            "base_kv" => g.base_kv[end],
            "bus_i" => g.bus_i[end]
        )
        push!(bus, [g.bus_i_str[end] bus_data])
    end

    local branch = []
    for g in groupby(df_branch_nordic, :name)
        branch_data = OrderedDict(
            "index" => g.name[end],
            "f_bus" => g.f_bus[end],
            "t_bus" => g.t_bus[end],
            "br_r" => g.br_r[end],
            "br_x" => g.br_x[end],
            "b_fr" => g.b_fr[end],
            "b_to" => g.b_to[end],
            "br_status" => 1,
            "base_kv" => g.v_nom[end],
            "c_rating_a" => g.c_rating_a[end]/1e3, # convert A to kA
            "angmin" => g.angmin[end],
            "angmax" => g.angmax[end],
            "transformer" => g.transformer[end],
            "tap" => g.tap[end],
            "shift" => g.shift[end],
            "g_fr" => g.g_fr[end],
            "g_to" => g.g_to[end],
            "tap" => g.tap[end]
        )
        push!(branch, [g.name_str[end] branch_data])
    end

    local dcline = []
    for g in groupby(df_links, :name)
        dcline_data = OrderedDict(
            "index" => g.name[end],
            "f_bus" => g.bus0[end],
            "t_bus" => g.bus1[end],
            "pminf" => -g.p_nom[end],
            "pmaxf" => g.p_nom[end],
            "pmint" => -g.p_nom[end],
            "pmaxt" => g.p_nom[end],
            "qminf" => 0,
            "qmaxf" => 0,
            "qmint" => 0,
            "qmaxt" => 0,
            "vmax" => 1.05,
            "loss0" => 0,
            "loss1" => 0,
            "vf" => 1.05,
            "vt" => 1.0,
            "br_status" => 1
        )
        push!(dcline, [g.name_str[end] dcline_data])
    end

    # we have to looop this methinks

    open("generated_jsons/nordic/nordic_500_h$hour.json", "w") do f
    end

    open("generated_jsons/nordic/nordic_500_h$hour.json", "a") do f
    write(f,"""{
        "baseMVA": 1000,
        "per_unit": false,
        "shunt":{},
        "storage":{},
        "switch":{},
        "dcline":{},
        "load":{""")
        JSON.print(f, load)
    write(f,    """},
        "gen":{""")
        JSON.print(f, gen)
    write(f,    """},
        "branch":{""")
        JSON.print(f, branch)
    write(f, """},
        "bus":{""")
        JSON.print(f, bus)
    write(f,    """}
    }""")
    end


    local test = read("generated_jsons//nordic/nordic_500_h$hour.json",String)
    local a = replace(test, "]],[[" => ",", "],[" => ":",  "]" => "", "[" => "")

    open("generated_jsons//nordic/nordic_500_h$hour.json", "w") do f
        write(f, a)
    end



    # # # # DK1 onlyyyyyyyy # # # # #

    # Using JSON.jl to make nested objects
    local load_DK1 = []
    local df_load = df_load2[in.(df_load2.load_bus, Ref(buses_DK1_all)), :] # only DK1
    for g in groupby(df_load, :load_bus)
        load_data = OrderedDict(
            "load_bus" => g.load_bus[end],
            "pd" => g.pd[end] * grid_loss_factor,
            "qd" => 0,
            "index" => g.load_bus[end],
            "status" => 1
        )
        push!(load_DK1, [g.load_bus_str[end] load_data])
    end

    local gen_DK1 = []
    local df_gen = df_gen[in.(df_gen.gen_bus, Ref(buses_DK1_all)), :] # only DK1
    for g in groupby(df_gen, :gen_bus)
        gen_data = OrderedDict(
            "gen_bus" => g.gen_bus[end],
            "pg" => g.pg[end],
            "qg" => 0,
            "pmax" => 9999,
            "pmin" => 0,
            "qmax" => 9999,
            "qmin" => -9999,
            "index" => g.gen_bus[end],
            "vg" => 1,
            "gen_status" => 1
        )
        push!(gen_DK1, [g.gen_bus_str[end] gen_data])
    end

    local bus_DK1 = []
    for g in groupby(df_bus_DK1, :bus_i)
        bus_data = OrderedDict(
            "bus_type" => g.bus_type[end], # 1 - PQ, 2 - PV, 3 - reference , 4 - isolated  , see https://matpower.org/docs/ref/matpower5.0/caseformat.html
            "vm" => 1,
            "va" => 0,
            "vmin" => 0.5,
            "vmax" => 1.5,
            "index" => g.bus_i[end],
            "gen_status" => 1,
            "base_kv" => g.base_kv[end],
            "bus_i" => g.bus_i[end]
        )
        push!(bus_DK1, [g.bus_i_str[end] bus_data])
    end

    local branch_DK1 = []
    for g in groupby(df_branch_DK1, :name)
        branch_data = OrderedDict(
            "index" => g.name[end],
            "f_bus" => g.f_bus[end],
            "t_bus" => g.t_bus[end],
            "br_r" => g.br_r[end],
            "br_x" => g.br_x[end],
            "b_fr" => g.b_fr[end],
            "b_to" => g.b_to[end],
            "br_status" => 1,
            "base_kv" => g.v_nom[end],
            "c_rating_a" => g.c_rating_a[end] /1e3, # convert A to kA
            "angmin" => g.angmin[end],
            "angmax" => g.angmax[end],
            "transformer" => g.transformer[end],
            "tap" => g.tap[end],
            "shift" => g.shift[end],
            "g_fr" => g.g_fr[end],
            "g_to" => g.g_to[end],
            "tap" => g.tap[end]
        )
        push!(branch_DK1, [g.name_str[end] branch_data])
    end



    open("generated_jsons//DK1/DK1_only_h$hour.json", "w") do f
    end

    open("generated_jsons//DK1/DK1_only_h$hour.json", "a") do f
    write(f,"""{
        "baseMVA": 1000,
        "per_unit": false,
        "shunt":{},
        "storage":{},
        "switch":{},
        "dcline":{},
        "load":{""")
        JSON.print(f, load_DK1)
    write(f,    """},
        "gen":{""")
        JSON.print(f, gen_DK1)
    write(f,    """},
        "branch":{""")
        JSON.print(f, branch_DK1)
    write(f, """},
        "bus":{""")
        JSON.print(f, bus_DK1)
    write(f,    """}
    }""")
    end


    local test = read("generated_jsons//DK1/DK1_only_h$hour.json",String)
    local a = replace(test, "]],[[" => ",", "],[" => ":",  "]" => "", "[" => "")

    open("generated_jsons//DK1/DK1_only_h$hour.json", "w") do f
        write(f, a)
    end

    if hour % 200 == 0
        println("Hour $hour")
    end
end

plot(imbalance, label = "imbalance")
plot(gens, label = "missing gen")
plot(loads, label = "missing load")