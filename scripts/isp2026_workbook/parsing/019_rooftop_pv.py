# %% [markdown]
# ## Rooftop PV
#
# Rooftop PV capacity and generation forecast.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Rooftop PV!B8:E10` (3 rows × 4 columns).

# %%
rooftop_pv_consultant_forecast_mapping = parse_spec('Rooftop PV', {'name': 'Consultant forecast mapping', 'range': 'B8:E10'})
show_table(rooftop_pv_consultant_forecast_mapping)

# %% [markdown]
# ### Rooftop PV capacity
#
# Source block: `Rooftop PV!B12:AH63` (52 rows × 33 columns).

# %%
rooftop_pv_rooftop_pv_capacity = parse_spec('Rooftop PV', {'name': 'Rooftop PV capacity', 'range': 'B12:AH63'})
show_table(rooftop_pv_rooftop_pv_capacity)

# %% [markdown]
# ### Rooftop PV energy
#
# Source block: `Rooftop PV!B65:AH116` (52 rows × 33 columns).

# %%
rooftop_pv_rooftop_pv_energy = parse_spec('Rooftop PV', {'name': 'Rooftop PV energy', 'range': 'B65:AH116'})
show_table(rooftop_pv_rooftop_pv_energy)

