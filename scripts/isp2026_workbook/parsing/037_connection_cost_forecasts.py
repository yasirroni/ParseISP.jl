# %% [markdown]
# ## Connection cost forecasts
#
# Forecast of transmission connection costs.

# %% [markdown]
# ### Wind and solar connection-cost forecasts
#
# Source block: `Connection cost forecasts!B8:AJ144` (137 rows × 35 columns).

# %%
connection_cost_forecasts_wind_and_solar_connection_cost_forecasts = parse_spec('Connection cost forecasts', {'name': 'Wind and solar connection-cost forecasts', 'range': 'B8:AJ144'})
show_table(connection_cost_forecasts_wind_and_solar_connection_cost_forecasts)

# %% [markdown]
# ### Other-generation connection-cost forecasts
#
# Source block: `Connection cost forecasts!B147:AJ388` (242 rows × 35 columns).

# %%
connection_cost_forecasts_other_generation_connection_cost_forecasts = parse_spec('Connection cost forecasts', {'name': 'Other-generation connection-cost forecasts', 'range': 'B147:AJ388'})
show_table(connection_cost_forecasts_other_generation_connection_cost_forecasts)

