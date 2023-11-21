using PowerModels
using Ipopt
using PowerModelsAnalytics


# importing the json file and parsing it to a network data dictionary
# this step also checks for obvious errors in the network (branches, buses etc that are not connected to anything)

hour = 1
region = "nordic"
#region = "DK1"
make_basic = "no"
PF = "DC"
type = "native"

# Importing the nordic system
if region == "nordic"
    network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/nordic/nordic_500_h$hour.json")
end
# importing the DK1 system
if region == "DK1"
    network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/DK1/DK1_only_h$hour.json")
end
#network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl//generated_jsons/DK1/DK1_only_h$hour.json")

#=display(network_data) # raw dictionary
PowerModels.print_summary(network_data) # quick table-like summary
PowerModels.make_mixed_units!(network_data)
PowerModels.print_summary(network_data) # quick table-like summary
PowerModels.make_per_unit!(network_data)
PowerModels.component_table(network_data, "bus", ["vmin", "vmax"]) # component data in matrix form =#

if make_basic == "yes"
    data = make_basic_network(network_data)
    else
    data = network_data
end




# Solving the DC powerflow problem
if PF == "DC"
    # Using the native solver
    dc_native_result = compute_dc_pf(data)
    PowerModels.update_data!(data, dc_native_result["solution"])
    # calculate dc branch flows from the solution
    branch_flows_dc = PowerModels.calc_branch_flow_dc(data)
end

# Using JuMP instead by making an optimization problem out of the DC powerflow problem (this can be more robust)
dc_pf_result = solve_pf(data, DCPPowerModel, Ipopt.Optimizer)
PowerModels.update_data!(data, dc_pf_result["solution"])

# Solving the AC powerflow problem
# Using the native solver
if PF == "AC"
    if type == "native"
        ac_native_result = compute_ac_pf(data)
        PowerModels.update_data!(data, ac_native_result["solution"])
        if ac_native_result["termination_status"] != true
            println("AC powerflow failed to converge at hour $hour")
        end
    else
        # Using JuMP instead by making an optimization problem out of the AC powerflow problem (this can be more robust)
        ac_pf_result = solve_pf(data, ACPPowerModel, Ipopt.Optimizer)
        PowerModels.update_data!(data, ac_pf_result["solution"])
    end
    # calculate ac branch flows from the solution
    branch_flows_ac = PowerModels.calc_branch_flow_ac(data)
    PowerModels.update_data!(data, branch_flows_ac)
end



# Table-like output of the solution
PowerModels.print_summary(ac_pf_result["solution"])
PowerModels.print_summary(ac_native_result["solution"])
PowerModels.print_summary(dc_native_result["solution"])
PowerModels.print_summary(branch_flows_ac)
PowerModels.print_summary(branch_flows_dc)

ac_native_result["solution"]["bus"]["7003"]

PowerModels.make_mixed_units!(data)
PowerModels.print_summary(data["bus"]) # quick table-like summary

export_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_json/$region/nordic_dc_h$hour.json", data)