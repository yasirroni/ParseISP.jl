# %% [markdown]
# ## Storage properties
#
# Battery storage to power ratio and round-trip efficiency.

# %% [markdown]
# ### Battery properties
#
# Source block: `Storage properties!B2:J19` (18 rows × 9 columns).

# %%
storage_properties_battery_properties = parse_spec('Storage properties', {'name': 'Battery properties', 'range': 'B2:J19'})
show_table(storage_properties_battery_properties)

# %% [markdown]
# ### Existing pumped-hydro properties
#
# Source block: `Storage properties!B21:E35` (15 rows × 4 columns).

# %%
storage_properties_existing_pumped_hydro_properties = parse_spec('Storage properties', {'name': 'Existing pumped-hydro properties', 'range': 'B21:E35'})
show_table(storage_properties_existing_pumped_hydro_properties)

# %% [markdown]
# ### New entrant pumped-hydro properties
#
# Source block: `Storage properties!G21:J27` (7 rows × 4 columns).

# %%
storage_properties_new_entrant_pumped_hydro_properties = parse_spec('Storage properties', {'name': 'New entrant pumped-hydro properties', 'range': 'G21:J27'})
show_table(storage_properties_new_entrant_pumped_hydro_properties)

# %% [markdown]
# ### On-site diesel storage
#
# Source block: `Storage properties!B38:C45` (8 rows × 2 columns).

# %%
storage_properties_on_site_diesel_storage = parse_spec('Storage properties', {'name': 'On-site diesel storage', 'range': 'B38:C45'})
show_table(storage_properties_on_site_diesel_storage)

