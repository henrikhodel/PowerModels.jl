# import packages/
using CSV, DataFrames, XLSX, Revise
using JSONTables, JSON, DataStructures, Plots
using PlotlyJS
using PowerModels

# Set the resolution (dpi) for the GR backend of Plots
gr(fmt = :tiff, size = (800, 400), dpi = 300)

# import which branches cross which BZ borders
df_branch_BZ = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/lines_crossing_BZ/Lines that cross BZ.csv", DataFrame)

df_bus_to_BZ = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/bus2BZ.csv", DataFrame)
df_border_to_BZ = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/border2BZ.csv", DataFrame)

# import import and export timeseries as dataframe in julia
df_internal_flows = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Grid data/Post_joint_fix/internal_flows.csv", DataFrame)

df_loads = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/input data/Scripts/entsoe_data/Processed data/Load_nordic_adjusted.csv", DataFrame)

# specify the hours to be processed
hours = 1:8760
# specify the voltage bounds
vm_min = 0.9875 # p.u.
vm_max = 1.05 # p.u.
# specify the maximum voltage angle difference
vad_max = 30 # degrees

# Initialize a new DataFrame to store the branch and bus results
df_flows = DataFrame(Hour = Int64[], Border = String[], Model_flow = Float64[], Power_to = Float64[], Power_from = Float64[], Physical_flow = Int64[])
df_losses = DataFrame(Hour = Int64[], Region = String[], Gen = Float64[], Load = Float64[], Net_Exp = Float64[], Losses = Float64[], Loss_percent = Float64[])

df_vm_violations = DataFrame(Hour = Int64[], Bus = Int64[], vm = Float64[])
df_vad_violations = DataFrame(Hour = Int64[], Branch = Int64[], vad = Float64[])
df_loading = DataFrame(Hour = Int64[], Branch = Int64[], loading = Float64[])
df_vad = DataFrame(Hour = Int64[], Branch = Int64[], vad = Float64[])

df_pg = DataFrame(Hour = Int64[], Bus = Int64[], pg = Float64[])
df_qg = DataFrame(Hour = Int64[], Bus = Int64[], qg = Float64[])
df_pg_qg_ratio = DataFrame(Bus = Int64[], max_pg = Float64[], max_qg = Float64[], min_qg = Float64[], max_qg_pg_ratio = Float64[], min_qg_pg_ratio = Float64[])

# Looping over the hours
for hour in hours
    # Importing the solved nordic system jsons as a dictionary
    data = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_json/nordic_h$hour.json")

    # Extracting the branch and bus results from the dictionary and putting them in a DataFrame
    df_branches = insertcols!(reduce(vcat, DataFrame.(values(data["branch_results"])), cols=:union), :Branch => parse.(Int,collect(keys(data["branch_results"]))))
    df_buses = insertcols!(reduce(vcat, DataFrame.(values(data["bus_results"])), cols=:union), :bus => parse.(Int,collect(keys(data["bus_results"]))))
    #df_buses_full = insertcols!(reduce(vcat, DataFrame.(values(data["bus"])), cols=:union), :bus => parse.(Int,collect(keys(data["bus_results"]))))

    power_sums = DataFrame(Border = String[], Model_flow = Float64[], Power_to = Float64[], Power_from = Float64[], Physical_flow = Int64[])

    # loss estimation for the nordic system
    BZ_loss = DataFrame(Region = String[], Gen = Float64[], Load = Float64[], Net_Exp = Float64[], Losses = Float64[], Loss_percent = Float64[])

    for row in eachrow(df_buses)
        bus_id = row.bus
        gen = row.pg
        load = row.pd
        vm = row.vm
        qgen = row.qg


        # Check if the voltage is within the specified bounds
        if vm < vm_min || vm > vm_max
            push!(df_vm_violations, (Hour = hour, Bus = bus_id, vm = vm))
        else   

        end

        push!(df_pg, (Hour = hour, Bus = bus_id, pg = gen))
        push!(df_qg, (Hour = hour, Bus = bus_id, qg = qgen))

        # Find the BZ for the current bus
        bus_BZ = df_bus_to_BZ[df_bus_to_BZ.bus .== bus_id, "BZ"]

        # if bus_BZ is not empty, extract the BZ name
        if !isempty(bus_BZ)
            bz = first(bus_BZ)
            
            # Check if the border already exists in power_sums
            idx = findfirst(BZ_loss.Region .== bz)
            
            if idx === nothing
                # If the border is not in power_sums, add it with the current power
                push!(BZ_loss, (Region = bz, Gen = gen, Load = load, Net_Exp = 0, Losses = 0, Loss_percent = 0))
            else
                # If the border is in power_sums, update the sum of power
                BZ_loss[idx, :Gen] += gen
                BZ_loss[idx, :Load] += load
            end
        end
    end

    ### sum the active power flow over branches crossing BZ borders
    # This is from Chat-GPT, don't judge me
    # Iterate through the branches in df_flows
    for row in eachrow(df_branches)
        branch_id = row.Branch
        power_to = row.p_to
        power_from = row.p_from
        vad = row.va_diff
        loading = row.loading
        if loading == nothing
            loading = 0
        end
        



        # Check if the voltage angle difference is within the specified bounds
        if abs(vad) > vad_max
            push!(df_vad_violations, (Hour = hour, Branch = branch_id, vad = vad))
        end

        if isfinite(vad)
            push!(df_vad, (Hour = hour, Branch = branch_id, vad = vad))
        end

        if isfinite(loading)
            if branch_id == 2017 # this 220kV line is probably triplex
                loading = loading / 1.5
            end
            push!(df_loading, (Hour = hour, Branch = branch_id, loading = loading))
        end

        # Find the border for the current branch
        border_row = df_branch_BZ[df_branch_BZ.Branch .== branch_id, "BZ crossing"]
        # find the factor for the current branch
        factor = df_branch_BZ[df_branch_BZ.Branch .== branch_id, "factor"]

        # if the border is not empty, extract the border name
        if !isempty(border_row)
            border = first(border_row)
            
            # Check if the border already exists in power_sums
            idx = findfirst(power_sums.Border .== border)
            model_flow = power_to .* -factor[1]
            
            if idx === nothing
                # If the border is not in power_sums, add it with the current power
                push!(power_sums, (Border = border, Model_flow = model_flow, Power_to = power_to, Power_from = power_from, Physical_flow = 0))
            else
                # If the border is in power_sums, update the sum of power
                power_sums[idx, :Power_to] += power_to
                power_sums[idx, :Power_from] += power_from
                power_sums[idx, :Model_flow] += model_flow
            end
        end
    end

    # map the physical flows to the power_sums dataframe
    for border in names(df_internal_flows[:,2:end])
        flow = df_internal_flows[hour, border]
        idx = findfirst(power_sums.Border .== border)
        power_sums[idx, :Physical_flow] += flow
    end

    # calculate the net export to the BZ
    for row in eachrow(df_border_to_BZ)
        factor = row.factor
        Border_id = row.Border
        BZ_id = row.BZ

        # find the net_exp of the border in power_sums
        net_exp = power_sums[power_sums.Border .== Border_id, :Model_flow] .* factor

        # find the index of the BZ in BZ_loss
        idx = findfirst(BZ_loss.Region .== BZ_id)
        BZ_loss[idx, :Net_Exp] += net_exp[1]
    end

    # Calculate the loss for each BZ
    for row in eachrow(BZ_loss)
        true_load = df_loads[hour, row.Region]
        idx = findfirst(BZ_loss.Region .== row.Region)
        BZ_loss[idx, :Losses] = BZ_loss[idx,:Gen] - BZ_loss[idx,:Load] - BZ_loss[idx,:Net_Exp]
        BZ_loss[idx, :Loss_percent] = BZ_loss[idx,:Losses] / true_load * 100 
    end

    # create a csv with the results data for each hour
    CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/nordic_branches_h$hour.csv", df_branches, transform=(col, val) -> something(val, missing))
    CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/nordic_buses_h$hour.csv", df_buses, transform=(col, val) -> something(val, missing))

    
    #print progress (so that I can relax and not worry that the code is stuck)
    if hour % 200 == 0
        println("Hour $hour")
    end

    insertcols!(BZ_loss, 1, :Hour => hour)
    # Push the rows for each hour into the new DataFrame
    for row in eachrow(BZ_loss)
        push!(df_losses, row)
    end

    insertcols!(power_sums, 1, :Hour => hour)
    # Push the rows from df1 into the new DataFrame
    for row in eachrow(power_sums)
        push!(df_flows, row)
    end

end

CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/losses.csv", df_losses, transform=(col, val) -> something(val, missing))
CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/flows.csv", df_flows, transform=(col, val) -> something(val, missing))
CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/vm_violations.csv", df_vm_violations, transform=(col, val) -> something(val, missing))
CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/vad_violations.csv", df_vad_violations, transform=(col, val) -> something(val, missing))

df_loading_violations = filter(row -> all(row.loading > 0.7), df_loading)
CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/loading_violations.csv", df_loading_violations, transform=(col, val) -> something(val, missing))


buses = unique(df_pg[:, :Bus])
for bus in buses
    idx = findall(df_pg.Bus .== bus)
    # find the maximum value of pg for the bus
    max_pg = maximum(df_pg[idx, :pg])
    # find the maximum value of qg for the bus
    idy = findall(df_qg.Bus .== bus)
    max_qg = maximum(df_qg[idy, :qg])
    min_qg = minimum(df_qg[idy, :qg])

    if max_pg == 0.0
        max_qg_pg_ratio = 0
        min_qg_pg_ratio = 0
    else
        max_qg_pg_ratio = max_qg / max_pg
        min_qg_pg_ratio = min_qg / max_pg
    end

    push!(df_pg_qg_ratio, (Bus = bus, max_pg = max_pg, max_qg = max_qg, min_qg = min_qg, max_qg_pg_ratio = max_qg / max_pg, min_qg_pg_ratio = min_qg / max_pg))
end

df_test = leftjoin(df_qg, df_pg_qg_ratio, on = :Bus)
df_test[!, :Result] = df_test.qg ./ df_test.max_pg

#CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/pg_qg_ratio_t.csv", df_test, transform=(col, val) -> something(val, missing))
CSV.write("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_csv/max_qg_pg_ratio.csv", df_pg_qg_ratio, transform=(col, val) -> something(val, missing))



df_test_plot = df_test[:, :Result]
filtered_data = [x for x in df_test_plot if isfinite(x)]
x = 1:length(filtered_data)
Plots.plot(x/length(x)*100, sort(filtered_data, rev = true), label = "qg / max_pg", title = "qg / max_pg", ylims = (-5,5))
Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/qg_ratio_max_pg.png")
# colour the line according to the pg_max, this way we can identify the buses that are "fake" gen buses i.e. they have very low pg_max and should not be classified as gen buses or at least have their reactive power capabiltites limited

df_test_plot = df_qg[:, :qg]
filtered_qg = [x for x in df_test_plot if isfinite(x)]
x = 1:length(filtered_data)
Plots.plot(x/length(x)*100, sort(filtered_data, rev = true), label = "qg", title = "qg", ylims = (-1000,300))
Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/qg_sorted.png")

df_plot = df_loading[:, :loading]
# remove the zeros (transformer loadings are zero)
df_plot = [x for x in df_plot if x > 0]
x = 1:length(df_plot)
df_plot = DataFrame(:loading => sort(df_plot, rev = true), :share => x/length(x)*100)
fig_plt = PlotlyJS.plot(
    df_plot, kind="scatter", mode="lines", x=:share, y=:loading, line=attr(width=3),
    labels = Dict(:share => "Share of hours [%]", :loading => "Line loading [%]"),
    Layout( 
        font_family = "Times New Roman",
        font_size = 16,
        font_color = "black",
        xaxis_range = (0,100),
        yaxis_range = (0,2),
        margin=attr(l=20,r=20,t=20,b=20)
    )
)
plt_str = "C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/loading_sorted.svg"
PlotlyJS.savefig(fig_plt, plt_str, width=round(Int, 300*1.5), height=300*1, scale=1)

df_plot = df_vad[:, :vad]
# remove the zeros (transformer loadings are zero)
df_plot = [abs(x) for x in df_plot]
x = 1:length(df_plot)
df_plot = DataFrame(:vad => sort(df_plot, rev = true), :share => x/length(x)*100)
fig_plt = PlotlyJS.plot(
    df_plot, kind="scatter", mode="lines", x=:share, y=:vad, line=attr(width=3),
    labels = Dict(:share => "Share of hours [%]", :vad => "Voltage angle difference [°]"),
    Layout( 
        font_family = "Times New Roman",
        font_size = 16,
        font_color = "black",
        xaxis_range = (0,100),
        yaxis_range = (0,45),
        margin=attr(l=20,r=20,t=20,b=20)
    )
)
plt_str = "C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/vad_sorted.svg"
PlotlyJS.savefig(fig_plt, plt_str, width=round(Int, 300*1.5), height=300*1, scale=1)

# generate timeseries plots for each border
for border in unique(df_flows[:,"Border"])
    idx = findall(df_flows.Border .== border)
    model = df_flows[idx, :Model_flow]
    physical = df_flows[idx, :Physical_flow]
    Plots.plot(df_flows.Hour[idx], [sort(model, rev = true), sort(physical, rev = true)], label = ["model" "physical"], title = border)
    xlabel!("Hour")
    ylabel!("Power [MW]")
    border = replace(border, ">" => "to")  # otherwise Windows will be angry >:(
    Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/$border duration.png")
end

for border in unique(df_flows[:,"Border"])
    idx = findall(df_flows.Border .== border)
    model = df_flows[idx, :Model_flow]
    physical = df_flows[idx, :Physical_flow]
    Plots.plot(df_flows.Hour[idx], [model, physical], label = ["model" "physical"], title = border)
    xlabel!("Hour")
    ylabel!("Power [MW]")
    border = replace(border, ">" => "to")  # otherwise Windows will be angry >:(
    Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/$border.png")
end


regions = unique(df_losses.Region)
i = 1
for region in regions
    idx = findall(df_losses.Region .== region)
    df_plot = df_losses[idx, :]
    if i == 1 
        Plots.plot(df_plot.Hour, df_plot.Loss_percent, label = region, ylims = (0,50))
    else
        Plots.plot!(df_plot.Hour, df_plot.Loss_percent, label = region, ylims = (0,50))
    end
    xlabel!("Hour")
    ylabel!("Losses [%]")
    title!("Transmission losses [%]")
    i = i+1
    Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/losses.png")
end

i = 1
for region in regions
    idx = findall(df_losses.Region .== region)
    df_plot = df_losses[idx, :]
    if i == 1 
        Plots.plot(df_plot.Hour, sort(df_plot.Loss_percent, rev = true), label = region, ylims = (0,50))
    else
        Plots.plot!(df_plot.Hour, sort(df_plot.Loss_percent, rev = true), label = region, ylims = (0,50))
    end
    xlabel!("Hour")
    ylabel!("Losses [%]")
    title!("Transmission losses [%]")
    i = i+1
    Plots.savefig("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/losses_duration.png")
end

test = parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_json/nordic/nordic_h154.json")


### Plot for the paper
df_residuals = CSV.read("C:/Users/hodel/OneDrive - Chalmers/PhD/Papers/Paper 2/Paper plots/residual_for_plots.csv", DataFrame)

# Create traces
trace1 = PlotlyJS.scatter(x=df_residuals[:,"hour"], y=df_residuals[:,"NO1 > NO2"],
                    mode="lines",
                    name="NO1 > NO2")
trace2 = PlotlyJS.scatter(x=df_residuals[:,"hour"], y=df_residuals[:,"NO2 > NO5"],
                    mode="lines",
                    name="NO2 > NO5",
                    marker_color="rgba(255, 0, 0, 0.5)"  # Set transparency to 50%
                    )
layout = Layout( 
    font_family = "Times New Roman",
    font_size = 16,
    font_color = "black",
    yaxis_range = (-600, 750),
    yaxis_title = "Residual [MW]",
    xaxis_title = "Hour of year [h]",
    xaxis_range = (0,8760),
    margin=attr(l=20,r=20,t=20,b=20),
    legend=attr(x=0.55, y=0.95)  # Adjust x and y to position the legend inside the plot
)

fig_plt = PlotlyJS.plot([trace1, trace2], layout)

plt_str = "C:/Users/hodel/Documents/GitHub/PowerModels.jl/result_plots/NO_res.svg"
PlotlyJS.savefig(fig_plt, plt_str, width=round(Int, 300*3), height=round(Int, 300*1.5), scale=1)