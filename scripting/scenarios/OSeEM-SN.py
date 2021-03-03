#Base Scenario Source Code

# Package Import
import os
import pandas as pd
import oemof.tabular.facades as fc
from oemof.tools.economics import annuity
import oemof.tabular.tools.postprocessing as pp
from oemof.solph import EnergySystem, Model, Bus

print("The necessary packages have been imported.")

# Input Directory
datapath = '/home/dozeummam/Insync/mnimaruf@gmail.com/Google Drive/projects/models/OSeEM-SN/data/base.xls'
#Modify this according to your own input datapath

# Results Directory
results_path = '/home/dozeummam/Insync/mnimaruf@gmail.com/Google Drive/projects/models/OSeEM-SN/results/scenarios/base'
if not os.path.exists(results_path):
    os.makedirs(results_path)
#Modify this according to your own result path

# Input Data Reading
timeseries = pd.read_excel(datapath, sheet_name="timeseries", index_col=[0], parse_dates=True)
timeseries.index.freq = "1H"
costs = pd.read_excel(datapath, sheet_name="costs", index_col=[0])
capacity = pd.read_excel(datapath, sheet_name="capacity", index_col=[0])

print("Data and results paths have been created. Input data have been read.")

# Energy System Creation
es = EnergySystem(timeindex=timeseries.index)
setattr(es, "typemap", fc.TYPEMAP)

print("The energy system has been created.")

# Bus Creation
elec_bus_SN = Bus(label="elec_Bus_SN")
heat_bus_SN = Bus(label="heat_Bus_SN")
fuel_bus_SN = Bus(label="fuel_Bus_SN")

# Bus Addition to the Energy Sytem
es.add(elec_bus_SN, heat_bus_SN, fuel_bus_SN)

print("The buses have been created")

# Energy System Components
# Onshore Wind
es.add(fc.Volatile(
    label="onshore_SN",
    bus=elec_bus_SN,
    carrier="wind",
    tech="onshore",
    profile=timeseries["onshore-SN"],
    expandable=True,
    capacity=capacity.at["sn_capacity", "onshore"],
    capacity_cost=(annuity(costs.at["capex", "onshore"], costs.at["lifetime", "onshore"], costs.at["wacc", "onshore"]))+(costs.at["fom", "onshore"]),
    marginal_cost=(costs.at["vom", "onshore"])+((costs.at["carrier_cost", "onshore"])/(capacity.at["efficiency", "onshore"])),
    capacity_potential=capacity.at["sn_capacity_potential", "onshore"]))

# Offshore Wind
es.add(fc.Volatile(
    label="offshore_SN",
    bus=elec_bus_SN,
    carrier="wind",
    tech="offshore",
    profile=timeseries["offshore"],
    expandable=True,
    capacity=capacity.at["sn_capacity", "offshore"],
    capacity_cost=(annuity(costs.at["capex", "offshore"], costs.at["lifetime", "offshore"], costs.at["wacc", "offshore"]))+(costs.at["fom", "offshore"]),
    marginal_cost=(costs.at["vom", "offshore"])+((costs.at["carrier_cost", "offshore"])/(capacity.at["efficiency", "offshore"])),
    capacity_potential=capacity.at["sn_capacity_potential", "offshore"]))

# Solar PV
es.add(fc.Volatile(
    label="pv_SN",
    bus=elec_bus_SN,
    carrier="solar",
    tech="pv",
    profile=timeseries["pv-SN"],
    expandable=True,
    capacity=capacity.at["sn_capacity", "pv"],
    capacity_cost=(annuity(costs.at["capex", "pv"], costs.at["lifetime", "pv"], costs.at["wacc", "pv"]))+(costs.at["fom", "pv"]),
    marginal_cost=(costs.at["vom", "pv"])+((costs.at["carrier_cost", "pv"])/(capacity.at["efficiency", "pv"])),
    capacity_potential=capacity.at["sn_capacity_potential", "pv"]))

# Hydro (Run-of-the-river)
es.add(fc.Volatile(
    label="hydro_SN",
    bus=elec_bus_SN,
    carrier="hydro",
    tech="ror",
    profile=timeseries["hydro-ror"],
    capacity=capacity.at["sn_capacity", "ror"],
    expandable=True,
    capacity_cost=(annuity(costs.at["capex", "ror"], costs.at["lifetime", "ror"], costs.at["wacc", "ror"]))+(costs.at["fom", "ror"]),
    marginal_cost=(costs.at["vom", "ror"])+((costs.at["carrier_cost", "ror"])/(capacity.at["efficiency", "ror"])),
    capacity_potential=capacity.at["sn_capacity_potential", "ror"]))

# Biomass-CHP
es.add(fc.Commodity(
    label='bm_comm_SN',
    bus=fuel_bus_SN,
    amount=capacity.at["amount", "sn_bm_commodity"],
    carrier='fuel',
    tech='commodity',
    marginal_cost=(costs.at["vom", "biomass"])+((costs.at["carrier_cost", "biomass"])/(capacity.at["efficiency", "biomass"]))))

es.add(fc.ExtractionTurbine(
    label="chp_SN",
    carrier='biomass',
    tech='ext',
    fuel_bus=fuel_bus_SN,
    heat_bus=heat_bus_SN,
    electricity_bus=elec_bus_SN,
    capacity=capacity.at["sn_capacity", "chp"],
    capacity_cost=(annuity(costs.at["capex", "biomass"], costs.at["lifetime", "biomass"], costs.at["wacc", "biomass"]))+(costs.at["fom", "biomass"]),
    carrier_cost=costs.at["carrier_cost", "biomass"],
    expandable=True,
    electric_efficiency=capacity.at["efficiency", "chp"],
    thermal_efficiency=capacity.at["efficiency", "chp"],
    condensing_efficiency=capacity.at["condensing_efficiency", "chp"],
    marginal_cost=(costs.at["vom", "biomass"])+((costs.at["carrier_cost", "biomass"])/(capacity.at["efficiency", "biomass"]))))

# Heat Pump
es.add(fc.Conversion(
    label='ashp_SN',
    from_bus=elec_bus_SN,
    to_bus=heat_bus_SN,
    carrier='electricity',
    tech='ashp',
    expandable=True,
    capacity=capacity.at["sn_capacity", "ashp"],
    capacity_cost=(annuity(costs.at["capex", "ashp"], costs.at["lifetime", "ashp"], costs.at["wacc", "ashp"]))+(costs.at["fom", "ashp"]),
    marginal_cost=(costs.at["vom", "ashp"])+((costs.at["carrier_cost", "ashp"])/(capacity.at["efficiency", "ashp"])),
    efficiency=capacity.at["efficiency", "ashp"]))

es.add(fc.Conversion(
    label='gshp_SN',
    from_bus=elec_bus_SN,
    to_bus=heat_bus_SN,
    carrier='electricity',
    tech='gshp',
    expandable=True,
    capacity=capacity.at["sn_capacity", "gshp"],
    capacity_cost=(annuity(costs.at["capex", "gshp"], costs.at["lifetime", "gshp"], costs.at["wacc", "gshp"]))+(costs.at["fom", "gshp"]),
    marginal_cost=(costs.at["vom", "gshp"])+((costs.at["carrier_cost", "gshp"])/(capacity.at["efficiency", "gshp"])),
    efficiency=capacity.at["efficiency", "gshp"]))

# Storage
# Pumped Hydro Storage
es.add(fc.Reservoir(
    label="phs_SN",
    bus=elec_bus_SN,
    carrier="hydro",
    tech="reservoir",
    capacity=capacity.at["sn_capacity", "phs"],
    storage_capacity=capacity.at["sn_storage_capacity", "phs"],
    efficiency=capacity.at["efficiency", "phs"],
    loss=capacity.at["loss", "phs"],
    marginal_cost=(costs.at["vom", "phs"])+((costs.at["carrier_cost", "phs"])/(capacity.at["efficiency", "phs"])),
    profile=timeseries["phs-SN"],
    initial_storage_level=0.5,
    balanced=True,
    invest_relation_input_capacity=1/(capacity.at["max_hours", "phs"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "phs"])))

# Li-ion Battery
es.add(fc.Storage(
    label="li-ion_SN",
    bus=elec_bus_SN,
    carrier="li-ion",
    tech="battery",
    expandable=True,
    loss=capacity.at["loss", "li-ion"],
    efficiency=capacity.at["efficiency", "li-ion"],
    storage_capacity=capacity.at["sn_storage_capacity", "li-ion"],
    storage_capacity_cost=annuity(costs.at["storage_capacity_cost", "li-ion"], costs.at["lifetime", "li-ion"], costs.at["wacc", "li-ion"])+(costs.at["fom", "li-ion"]),
    storage_capacity_potential=capacity.at["sn_storage_capacity_potential", "li-ion"],
    capacity=capacity.at["sn_capacity", "li-ion"],
    capacity_cost=(annuity(costs.at["capex", "li-ion"], costs.at["lifetime", "li-ion"], costs.at["wacc", "li-ion"])),
    marginal_cost=(costs.at["vom", "li-ion"])+((costs.at["carrier_cost", "li-ion"])/(capacity.at["efficiency", "li-ion"])),
    capacity_potential=capacity.at["sn_capacity_potential", "li-ion"],
    invest_relation_input_capacity=1/(capacity.at["max_hours", "li-ion"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "li-ion"])))

# Adiabatic Compressed Air Energy Storage
es.add(fc.Storage(
    label="acaes_SN",
    bus=elec_bus_SN,
    carrier="cavern",
    tech="acaes",
    expandable=True,
    loss=capacity.at["loss", "acaes"],
    efficiency=capacity.at["efficiency", "acaes"],
    storage_capacity=capacity.at["sn_storage_capacity", "acaes"],
    storage_capacity_cost=annuity(costs.at["storage_capacity_cost", "acaes"], costs.at["lifetime", "acaes"], costs.at["wacc", "acaes"])+(costs.at["fom", "acaes"]),
    storage_capacity_potential=capacity.at["sn_storage_capacity_potential", "acaes"],
    capacity=capacity.at["sn_capacity", "acaes"],
    capacity_cost=annuity(costs.at["capex", "acaes"], costs.at["lifetime", "acaes"], costs.at["wacc", "acaes"]),
    marginal_cost=(costs.at["vom", "acaes"])+((costs.at["carrier_cost", "acaes"])/(capacity.at["efficiency", "acaes"])),
    capacity_potential=capacity.at["sn_capacity_potential", "acaes"],
    invest_relation_input_capacity=1/(capacity.at["max_hours", "acaes"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "acaes"])))

# Redox Flow Battery
es.add(fc.Storage(
    label="redox_SN",
    bus=elec_bus_SN,
    carrier="redox",
    tech="battery",
    expandable=True,
    loss=capacity.at["loss", "redox"],
    efficiency=capacity.at["efficiency", "redox"],
    storage_capacity=capacity.at["sn_storage_capacity", "redox"],
    storage_capacity_cost=annuity(costs.at["storage_capacity_cost", "redox"], costs.at["lifetime", "redox"], costs.at["wacc", "redox"])+(costs.at["fom", "redox"]),
    storage_capacity_potential=capacity.at["sn_storage_capacity_potential", "redox"],
    capacity=capacity.at["sn_capacity", "redox"],
    capacity_cost=annuity(costs.at["capex", "redox"], costs.at["lifetime", "redox"], costs.at["wacc", "redox"]),
    marginal_cost=(costs.at["vom", "redox"])+((costs.at["carrier_cost", "redox"])/(capacity.at["efficiency", "redox"])),
    capacity_potential=capacity.at["sn_capacity_potential", "redox"],
    invest_relation_input_capacity=1/(capacity.at["max_hours", "redox"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "redox"])))

# Hydrogen
es.add(fc.Storage(
    label="hydrogen_SN",
    bus=elec_bus_SN,
    carrier="hydrogen",
    tech="hydrogen",
    expandable=True,
    loss=capacity.at["loss", "hydrogen"],
    efficiency=capacity.at["efficiency", "hydrogen"],
    storage_capacity=capacity.at["sn_storage_capacity", "hydrogen"],
    storage_capacity_cost=annuity(costs.at["storage_capacity_cost", "hydrogen"], costs.at["lifetime", "hydrogen"], costs.at["wacc", "hydrogen"])+(costs.at["fom", "hydrogen"]),
    storage_capacity_potential=capacity.at["sn_storage_capacity_potential", "hydrogen"],
    capacity=capacity.at["sn_capacity", "hydrogen"],
    capacity_cost=annuity(costs.at["capex", "hydrogen"], costs.at["lifetime", "hydrogen"], costs.at["wacc", "hydrogen"]),
    marginal_cost=(costs.at["vom", "hydrogen"])+((costs.at["carrier_cost", "hydrogen"])/(capacity.at["efficiency", "hydrogen"])),
    capacity_potential=capacity.at["sn_capacity_potential", "hydrogen"],
    invest_relation_input_capacity=1/(capacity.at["max_hours", "hydrogen"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "hydrogen"])))

# Thermal Energy Storage - Hot Water Tanks
es.add(fc.Storage(
    label="tes_SN",
    bus=heat_bus_SN,
    carrier="heat",
    tech="tes",
    expandable=True,
    loss=capacity.at["loss", "tes"],
    efficiency=capacity.at["efficiency", "tes"],
    storage_capacity=capacity.at["sn_storage_capacity", "tes"],
    storage_capacity_cost=annuity(costs.at["storage_capacity_cost", "tes"], costs.at["lifetime", "tes"], costs.at["wacc", "tes"])+(costs.at["fom", "tes"]),
    storage_capacity_potential=capacity.at["sn_storage_capacity_potential", "tes"],
    capacity=capacity.at["sn_capacity", "tes"],
    capacity_cost=annuity(costs.at["capex", "tes"], costs.at["lifetime", "tes"], costs.at["wacc", "tes"]),
    marginal_cost=(costs.at["vom", "tes"])+((costs.at["carrier_cost", "tes"])/(capacity.at["efficiency", "tes"])),
    capacity_potential=capacity.at["sn_capacity_potential", "tes"],
    invest_relation_input_capacity=1/(capacity.at["max_hours", "tes"]),
    invest_relation_output_capacity=1/(capacity.at["max_hours", "tes"])))

print("All components have been added to the energy system.")

# Demand Data
# Electricity
es.add(fc.Load(
    label="electricity_SN",
    bus=elec_bus_SN,
    carrier='electricity',
    amount=capacity.at["amount", "sn_electricity"],
    profile=timeseries["electricity-SN"]))

# Heat
es.add(fc.Load(
    label="space_SN",
    bus=heat_bus_SN,
    carrier='heat',
    amount=capacity.at["amount", "sn_space_heat"],
    profile=timeseries["space-SN"]))

es.add(fc.Load(
    label="dhw_SN",
    bus=heat_bus_SN,
    carrier='heat',
    amount=capacity.at["amount", "sn_dhw_heat"],
    profile=timeseries["dhw-SN"]))

# Excess
es.add(fc.Excess(
    label="excess_SN",
    bus=elec_bus_SN,
    tech="grid",
    carrier="electricity",
    marginal_cost=costs.at["vom", "excess"]))

es.add(fc.Excess(
    label="excess_heat_SN",
    bus=heat_bus_SN,
    tech="grid",
    carrier="heat",
    marginal_cost=costs.at["vom", "excess_heat"]))

print("Demand data have been read.")

# OEMoF Model Creation
m = Model(es)

print("OSeEM-SN is ready to solve.")

# LP File
m.write(os.path.join(results_path, "investment.lp"), io_options={"symbolic_solver_labels": True})

# Shadow Price
m.receive_duals()

# Solve
m.solve("cbc")
m.results = m.results()

print("OSeEM-SN solved the optimization problem. :)")

# Results
pp.write_results(m, results_path)

print("Results have been written. Results are available in {}.".format(results_path))
