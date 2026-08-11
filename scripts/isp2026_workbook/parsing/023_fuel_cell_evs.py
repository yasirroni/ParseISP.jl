# %% [markdown]
# ## Fuel cell EVs
#
# Fuel cell electric vehicles uptake.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Fuel cell EVs!B7:E9` (3 rows × 4 columns).

# %%
fuel_cell_evs_consultant_forecast_mapping = parse_spec('Fuel cell EVs', {'name': 'Consultant forecast mapping', 'range': 'B7:E9'})
show_table(fuel_cell_evs_consultant_forecast_mapping)

# %% [markdown]
# ### Fuel-cell EV uptake
#
# Source block: `Fuel cell EVs!B11:AH62` (52 rows × 33 columns).

# %%
fuel_cell_evs_fuel_cell_ev_uptake = parse_spec('Fuel cell EVs', {'name': 'Fuel-cell EV uptake', 'range': 'B11:AH62'})
show_table(fuel_cell_evs_fuel_cell_ev_uptake)

