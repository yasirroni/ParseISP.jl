# %% [markdown]
# ## Data Centre Forecasts
#
# Forecast of electricity consumption from data centre growth.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Data Centre Forecasts!B6:E8` (3 rows × 4 columns).

# %%
data_centre_forecasts_consultant_forecast_mapping = parse_spec('Data Centre Forecasts', {'name': 'Consultant forecast mapping', 'range': 'B6:E8'})
show_table(data_centre_forecasts_consultant_forecast_mapping)

# %% [markdown]
# ### Data-centre demand — Slower Growth
#
# Source block: `Data Centre Forecasts!B10:AF16` (7 rows × 31 columns).

# %%
data_centre_forecasts_data_centre_demand_slower_growth = parse_spec('Data Centre Forecasts', {'name': 'Data-centre demand — Slower Growth', 'range': 'B10:AF16'})
show_table(data_centre_forecasts_data_centre_demand_slower_growth)

# %% [markdown]
# ### Data-centre demand — Step Change
#
# Source block: `Data Centre Forecasts!B18:AF24` (7 rows × 31 columns).

# %%
data_centre_forecasts_data_centre_demand_step_change = parse_spec('Data Centre Forecasts', {'name': 'Data-centre demand — Step Change', 'range': 'B18:AF24'})
show_table(data_centre_forecasts_data_centre_demand_step_change)

# %% [markdown]
# ### Data-centre demand — Accelerated Transition
#
# Source block: `Data Centre Forecasts!B26:AF32` (7 rows × 31 columns).

# %%
data_centre_forecasts_data_centre_demand_accelerated_transition = parse_spec('Data Centre Forecasts', {'name': 'Data-centre demand — Accelerated Transition', 'range': 'B26:AF32'})
show_table(data_centre_forecasts_data_centre_demand_accelerated_transition)

