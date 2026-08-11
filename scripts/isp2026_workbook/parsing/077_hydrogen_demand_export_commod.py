# %% [markdown]
# ## Hydrogen demand-Export&Commod
#
# Projections of demand for export hydrogen and green commodities production.

# %% [markdown]
# ### Hydrogen export demand
#
# Source block: `Hydrogen demand-Export&Commod!B2:AH52` (51 rows × 33 columns).

# %%
hydrogen_demand_export_and_commod_hydrogen_export_demand = parse_spec('Hydrogen demand-Export&Commod', {'name': 'Hydrogen export demand', 'range': 'B2:AH52'})
show_table(hydrogen_demand_export_and_commod_hydrogen_export_demand)

# %% [markdown]
# ### Hydrogen demand for green commodities
#
# Source block: `Hydrogen demand-Export&Commod!B54:AH105` (52 rows × 33 columns).

# %%
hydrogen_demand_export_and_commod_hydrogen_demand_for_green_commodities = parse_spec('Hydrogen demand-Export&Commod', {'name': 'Hydrogen demand for green commodities', 'range': 'B54:AH105'})
show_table(hydrogen_demand_export_and_commod_hydrogen_demand_for_green_commodities)

# %% [markdown]
# ### Electricity demand for green steel
#
# Source block: `Hydrogen demand-Export&Commod!B107:AH156` (50 rows × 33 columns).

# %%
hydrogen_demand_export_and_commod_electricity_demand_for_green_steel = parse_spec('Hydrogen demand-Export&Commod', {'name': 'Electricity demand for green steel', 'range': 'B107:AH156'})
show_table(hydrogen_demand_export_and_commod_electricity_demand_for_green_steel)

