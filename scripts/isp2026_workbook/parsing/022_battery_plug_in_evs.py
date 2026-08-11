# %% [markdown]
# ## Battery & Plug-in EVs
#
# Battery and plug-in electric vehicles uptake and energy consumption for driving purposes.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Battery & Plug-in EVs!B7:E9` (3 rows × 4 columns).

# %%
battery_and_plug_in_evs_consultant_forecast_mapping = parse_spec('Battery & Plug-in EVs', {'name': 'Consultant forecast mapping', 'range': 'B7:E9'})
show_table(battery_and_plug_in_evs_consultant_forecast_mapping)

# %% [markdown]
# ### BEV and PHEV energy
#
# Source block: `Battery & Plug-in EVs!B11:AH62` (52 rows × 33 columns).

# %%
battery_and_plug_in_evs_bev_and_phev_energy = parse_spec('Battery & Plug-in EVs', {'name': 'BEV and PHEV energy', 'range': 'B11:AH62'})
show_table(battery_and_plug_in_evs_bev_and_phev_energy)

# %% [markdown]
# ### BEV and PHEV uptake
#
# Source block: `Battery & Plug-in EVs!B64:AH115` (52 rows × 33 columns).

# %%
battery_and_plug_in_evs_bev_and_phev_uptake = parse_spec('Battery & Plug-in EVs', {'name': 'BEV and PHEV uptake', 'range': 'B64:AH115'})
show_table(battery_and_plug_in_evs_bev_and_phev_uptake)

