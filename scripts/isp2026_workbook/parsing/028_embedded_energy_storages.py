# %% [markdown]
# ## Embedded energy storages
#
# Embedded consumer energy storage (battery) forecast.

# %% [markdown]
# ### Forecast mapping
#
# Source block: `Embedded energy storages!B7:E9` (3 rows × 4 columns).

# %%
embedded_energy_storages_forecast_mapping = parse_spec('Embedded energy storages', {'name': 'Forecast mapping', 'range': 'B7:E9'})
show_table(embedded_energy_storages_forecast_mapping)

# %% [markdown]
# ### Embedded energy storage capacity
#
# Source block: `Embedded energy storages!B11:AH62` (52 rows × 33 columns).

# %%
embedded_energy_storages_embedded_energy_storage_capacity = parse_spec('Embedded energy storages', {'name': 'Embedded energy storage capacity', 'range': 'B11:AH62'})
show_table(embedded_energy_storages_embedded_energy_storage_capacity)

# %% [markdown]
# ### Embedded energy storage degraded energy
#
# Source block: `Embedded energy storages!B65:AH116` (52 rows × 33 columns).

# %%
embedded_energy_storages_embedded_energy_storage_degraded_energy = parse_spec('Embedded energy storages', {'name': 'Embedded energy storage degraded energy', 'range': 'B65:AH116'})
show_table(embedded_energy_storages_embedded_energy_storage_degraded_energy)

