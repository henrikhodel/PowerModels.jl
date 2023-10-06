using PowerModels
using Ipopt
using Suppressor


# importing the json file and parsing it to a network data dictionary
# this step also checks for obvious errors in the network (branches, buses etc that are not connected to anything)
hours = 1:8760

#region = "nordic"
region = "DK1"
make_basic = "no"
PF = "AC"
type = "native"
make_basic = "no"
Solutions = []
silence()
for hour in hours
    #println(hour)
    # Importing the nordic system
    if region == "nordic"
        network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/nordic/nordic_500_h$hour.json")
    end
    # importing the DK1 system
    if region == "DK1"
        network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/DK1/DK1_only_h$hour.json")
    end

    # Making the network basic
    if make_basic == "yes"
        data = make_basic_network(network_data)
        else
        data = network_data
    end


    # Solving the DC powerflow problem
    # Using the native solver
    if PF == "DC"
        dc_native_result = compute_dc_pf(data)
        PowerModels.update_data!(data, dc_native_result["solution"])

        # Using JuMP instead by making an optimization problem out of the DC powerflow problem (this can be more robust)
        #dc_pf_result = solve_pf(data, DCPPowerModel, Ipopt.Optimizer)
        #PowerModels.update_data!(data, dc_pf_result["solution"])

        # calculate dc branch flows from the solution
        #branch_flows_dc = PowerModels.calc_branch_flow_dc(data)
    end

    # Solving the AC powerflow problem
    # Using the native solver
    if PF == "AC"
        if type == "native"
            try
                ac_native_result = compute_ac_pf(data)
                PowerModels.update_data!(data, ac_native_result["solution"])
                branch_flows_ac = PowerModels.calc_branch_flow_ac(data)
                if ac_native_result["termination_status"] != true
                    println("AC powerflow failed to converge at hour $hour")
                end
            catch
                println("AC powerflow failed to be solved at hour $hour")
            end
        else
            try
                # Using JuMP instead by making an optimization problem out of the AC powerflow problem (this can be more robust)
                ac_pf_result = solve_pf(data, ACPPowerModel, Ipopt.Optimizer)
                PowerModels.update_data!(data, ac_pf_result["solution"])
                branch_flows_ac = PowerModels.calc_branch_flow_ac(data)
            catch
                println("AC powerflow failed to be solved at hour $hour")
            end
        end
        # calculate ac branch flows from the solution
        
    end

    #push!(Solutions, [hour data])   

end


# Plotting etc

#plot_network(network_data)

# Table-like output of the solution
#PowerModels.print_summary(ac_pf_result["solution"])
#PowerModels.print_summary(ac_native_result["solution"])
#PowerModels.print_summary(dc_native_result["solution"])
#PowerModels.print_summary(branch_flows_ac)
#PowerModels.print_summary(branch_flows_dc)

#result["solution"]["bus"]["1"]

#PowerModels.make_mixed_units!(data)
#PowerModels.print_summary(data) # quick table-like summary