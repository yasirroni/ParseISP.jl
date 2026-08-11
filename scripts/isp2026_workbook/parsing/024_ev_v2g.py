# %% [markdown]
# ## EV V2G
#
# Vehicle to Grid battery characteristics.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `EV V2G!B8:E10` (3 rows × 4 columns).

# %%
ev_v2g_consultant_forecast_mapping = parse_spec('EV V2G', {'name': 'Consultant forecast mapping', 'range': 'B8:E10'})
show_table(ev_v2g_consultant_forecast_mapping)

# %% [markdown]
# ### Vehicle-to-grid capacity
#
# Source block: `EV V2G!B12:AH62` (51 rows × 33 columns).

# %%
ev_v2g_vehicle_to_grid_capacity = parse_spec('EV V2G', {'name': 'Vehicle-to-grid capacity', 'range': 'B12:AH62'})
show_table(ev_v2g_vehicle_to_grid_capacity)

# %% [markdown]
# ### Vehicle-to-grid depth
#
# Source block: `EV V2G!B64:AH115` (52 rows × 33 columns).

# %%
ev_v2g_vehicle_to_grid_depth = parse_spec('EV V2G', {'name': 'Vehicle-to-grid depth', 'range': 'B64:AH115'})
show_table(ev_v2g_vehicle_to_grid_depth)

