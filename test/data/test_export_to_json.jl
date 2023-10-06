using PowerModels

# import matpower file
network_data = PowerModels.parse_file("C:/Users/hodel/Documents/GitHub/PowerModels.jl/test/data/matpower/case5_dc.m")

correct_network_data!

# export network data to json
PowerModels.export_file("case5_dc.json", network_data)
