using Revise
using PowerModels
using Ipopt


# importing the json file and parsing it to a network data dictionary
# this step also checks for obvious errors in the network (branches, buses etc that are not connected to anything)
hours = 1:8760

PF = "AC"
type = "native"
make_basic = "no"
Solutions = []
silence()
for hour in hours
    #println(hour)
    # Importing the nordic system
    network_data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/generated_jsons/nordic/nordic_500_h$hour.json")

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
                PowerModels.update_data!(data, branch_flows_ac)
                PowerModels.make_mixed_units!(data)
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
                PowerModels.update_data!(data, branch_flows_ac)
                PowerModels.make_mixed_units!(data)
            catch
                println("AC powerflow failed to be solved at hour $hour")
            end
        end
        # extract the bus data that we are interested in
        vm = Dict(bus["index"] => bus["vm"] for (i,bus) in data["bus"])
        va = Dict(bus["index"] => bus["va"] for (i,bus) in data["bus"])
        pd = Dict(load["index"] => load["pd"] for (i,load) in data["load"])
        pg = Dict(gen["index"] => gen["pg"] for (i,gen) in data["gen"])
        qg = Dict(gen["index"] => gen["qg"] for (i,gen) in data["gen"])
        #println(data["load"])
        bus_results = Dict{String,Any}()
        for (i,bus) in data["bus"]
            vm = bus["vm"]
            va = bus["va"]
            pd2 = 0
            pg2 = 0
            qg2 = 0
            try pd2 = pd[parse.(Int,i)]
            catch 
                pd2 = 0
            end
            try pg2 = pg[parse.(Int,i)]
            catch 
                pg2 = 0
            end
            try qg2 = qg[parse.(Int,i)]
            catch 
                qg2 = 0
            end

            bus_results[i] = Dict(
                "vm" => vm,
                "va" => va,
                "pd" => pd2,
                "pg" => pg2,
                "qg" => qg2
            )
        end

        # calculate loading of lines
        loading = Dict(branch["index"] => branch["loading"] for (i,branch) in data["branch"])
        S_rating = Dict(branch["index"] => branch["S_rate"] for (i,branch) in data["branch"])
        S_from = Dict(branch["index"] => branch["Sf"] for (i,branch) in data["branch"])
        S_to = Dict(branch["index"] => branch["St"] for (i,branch) in data["branch"])
        P_to = Dict(branch["index"] => branch["pt"] for (i,branch) in data["branch"])
        P_from = Dict(branch["index"] => branch["pf"] for (i,branch) in data["branch"])
        vad = Dict(branch["index"] => branch["vad"] for (i,branch) in data["branch"])
        
        branch_results = Dict{String,Any}()
        for (i,branch) in data["branch"]
            loading = branch["loading"]
            S_rating = branch["S_rate"]
            S_from = branch["Sf"]
            S_to = branch["St"]
            P_to = branch["pt"]
            P_from = branch["pf"]
            vad = branch["vad"]
            
            branch_results[i] = Dict(
                "loading" => loading,
                "S_rating" => S_rating,
                "S_from" => S_from,
                "S_to" => S_to,
                "p_to" => P_to,
                "p_from" => P_from,
                "va_diff" => vad
            )
        end

        PowerModels.update_data!(data, Dict{String,Any}("branch_results" => branch_results))
        PowerModels.update_data!(data, Dict{String,Any}("bus_results" => bus_results))
    end

    export_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_json/nordic_h$hour.json", data)

    if hour % 200 == 0
        println("Hour $hour")
    end
    #push!(Solutions, [data_export])   
end
println("Successfully exported results")