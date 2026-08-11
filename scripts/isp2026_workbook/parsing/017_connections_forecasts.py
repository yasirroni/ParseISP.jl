# %% [markdown]
# ## Connections Forecasts
#
# Residential connections forecasts.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Connections Forecasts!B8:E11` (4 rows × 4 columns).

# %%
connections_forecasts_consultant_forecast_mapping = parse_spec('Connections Forecasts', {'name': 'Consultant forecast mapping', 'range': 'B8:E11'})
show_table(connections_forecasts_consultant_forecast_mapping)

# %% [markdown]
# ### Residential connections — Slower Growth
#
# Source block: `Connections Forecasts!B15:AH22` (8 rows × 33 columns).

# %%
connections_forecasts_residential_connections_slower_growth = parse_spec('Connections Forecasts', {'name': 'Residential connections — Slower Growth', 'range': 'B15:AH22'})
show_table(connections_forecasts_residential_connections_slower_growth)

# %% [markdown]
# ### Residential connections — Step Change
#
# Source block: `Connections Forecasts!B24:AH31` (8 rows × 33 columns).

# %%
connections_forecasts_residential_connections_step_change = parse_spec('Connections Forecasts', {'name': 'Residential connections — Step Change', 'range': 'B24:AH31'})
show_table(connections_forecasts_residential_connections_step_change)

# %% [markdown]
# ### Residential connections — Accelerated Transition
#
# Source block: `Connections Forecasts!B33:AH40` (8 rows × 33 columns).

# %%
connections_forecasts_residential_connections_accelerated_transition = parse_spec('Connections Forecasts', {'name': 'Residential connections — Accelerated Transition', 'range': 'B33:AH40'})
show_table(connections_forecasts_residential_connections_accelerated_transition)

