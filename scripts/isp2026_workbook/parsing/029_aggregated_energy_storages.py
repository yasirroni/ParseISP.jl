# %% [markdown]
# ## Aggregated energy storages
#
# The 'aggregated' share of embedded energy storages that is modelled like a Virtual Power Plant (VPP).

# %% [markdown]
# ### Forecast mapping
#
# Source block: `Aggregated energy storages!B7:E9` (3 rows × 4 columns).

# %%
aggregated_energy_storages_forecast_mapping = parse_spec('Aggregated energy storages', {'name': 'Forecast mapping', 'range': 'B7:E9'})
show_table(aggregated_energy_storages_forecast_mapping)

# %% [markdown]
# ### Aggregated energy storage capacity
#
# Source block: `Aggregated energy storages!B11:AH62` (52 rows × 33 columns).

# %%
aggregated_energy_storages_aggregated_energy_storage_capacity = parse_spec('Aggregated energy storages', {'name': 'Aggregated energy storage capacity', 'range': 'B11:AH62'})
show_table(aggregated_energy_storages_aggregated_energy_storage_capacity)

# %% [markdown]
# ### Aggregated energy storage degraded energy
#
# Source block: `Aggregated energy storages!B65:AH116` (52 rows × 33 columns).

# %%
aggregated_energy_storages_aggregated_energy_storage_degraded_energy = parse_spec('Aggregated energy storages', {'name': 'Aggregated energy storage degraded energy', 'range': 'B65:AH116'})
show_table(aggregated_energy_storages_aggregated_energy_storage_degraded_energy)

