# %% [markdown]
# ## PVNSG
#
# PV non-scheduled generation (PVNSG) capacity and generation forecast.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `PVNSG!B8:E10` (3 rows × 4 columns).

# %%
pvnsg_consultant_forecast_mapping = parse_spec('PVNSG', {'name': 'Consultant forecast mapping', 'range': 'B8:E10'})
show_table(pvnsg_consultant_forecast_mapping)

# %% [markdown]
# ### PVNSG capacity
#
# Source block: `PVNSG!B12:AH63` (52 rows × 33 columns).

# %%
pvnsg_pvnsg_capacity = parse_spec('PVNSG', {'name': 'PVNSG capacity', 'range': 'B12:AH63'})
show_table(pvnsg_pvnsg_capacity)

# %% [markdown]
# ### PVNSG energy
#
# Source block: `PVNSG!B65:AH116` (52 rows × 33 columns).

# %%
pvnsg_pvnsg_energy = parse_spec('PVNSG', {'name': 'PVNSG energy', 'range': 'B65:AH116'})
show_table(pvnsg_pvnsg_energy)

