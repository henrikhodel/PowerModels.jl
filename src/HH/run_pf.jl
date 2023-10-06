using PowerModels
using Ipopt
using PowerModelsAnalytics


# importing the json file and parsing it to a network data dictionary
# this step also checks for obvious errors in the network (branches, buses etc that are not connected to anything)

hour = 1
#region = "nordic"
region = "DK1"
make_basic = "no"
PF = "AC"

# Importing the nordic system
if region == "nordic"
    network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/nordic/nordic_500_h$hour.json")
end
# importing the DK1 system
if region == "DK1"
    network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/DK1/DK1_only_h$hour.json")
end
#network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl//generated_jsons/DK1/DK1_only_h$hour.json")

display(network_data) # raw dictionary
PowerModels.print_summary(network_data) # quick table-like summary
PowerModels.make_mixed_units!(network_data)
PowerModels.print_summary(network_data) # quick table-like summary
PowerModels.make_per_unit!(network_data)
PowerModels.component_table(network_data, "bus", ["vmin", "vmax"]) # component data in matrix form

#test_am = calc_admittance_matrix(network_data)

data = make_basic_network(network_data)
data = network_data
#export_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/basic_nordic_system.json", test_data)


#PowerModels.connected_components(data)
# Solving the AC powerflow problem



# Solving the DC powerflow problem
# Using the native solver
dc_native_result = compute_dc_pf(data)
PowerModels.update_data!(data, dc_native_result["solution"])

# Using JuMP instead by making an optimization problem out of the DC powerflow problem (this can be more robust)
dc_pf_result = solve_pf(data, DCPPowerModel, Ipopt.Optimizer)
PowerModels.update_data!(data, dc_pf_result["solution"])

# Solving the AC powerflow problem
# Using the native solver
ac_native_result = compute_ac_pf(data, show_trace=true)
PowerModels.update_data!(data, ac_native_result["solution"])


# Using JuMP instead by making an optimization problem out of the AC powerflow problem (this can be more robust)
ac_pf_result = solve_pf(data, ACPPowerModel, Ipopt.Optimizer)
PowerModels.update_data!(data, ac_pf_result["solution"])




# calculate ac branch flows from the solution
branch_flows_ac = PowerModels.calc_branch_flow_ac(data)

# calculate dc branch flows from the solution
branch_flows_dc = PowerModels.calc_branch_flow_dc(data)

# Plotting etc

#plot_network(network_data)

# Table-like output of the solution
PowerModels.print_summary(ac_pf_result["solution"])
PowerModels.print_summary(ac_native_result["solution"])
PowerModels.print_summary(dc_native_result["solution"])
PowerModels.print_summary(branch_flows_ac)
PowerModels.print_summary(branch_flows_dc)

result["solution"]["bus"]["1"]

PowerModels.make_mixed_units!(data)
PowerModels.print_summary(data) # quick table-like summary

