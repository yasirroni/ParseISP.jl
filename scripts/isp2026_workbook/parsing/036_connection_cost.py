# %% [markdown]
# ## Connection cost
#
# Cost to connect different generation technologies.

# %% [markdown]
# ### Wind and solar connection costs
#
# Source block: `Connection cost!B6:J61` (56 rows × 9 columns).

# %%
connection_cost_wind_and_solar_connection_costs = parse_spec('Connection cost', {'name': 'Wind and solar connection costs', 'range': 'B6:J61'})
show_table(connection_cost_wind_and_solar_connection_costs)

# %% [markdown]
# ### Other-generation regional connection costs
#
# Source block: `Connection cost!B62:R73` (12 rows × 17 columns).

# %%
connection_cost_other_generation_regional_connection_costs = parse_spec('Connection cost', {'name': 'Other-generation regional connection costs', 'range': 'B62:R73'})
show_table(connection_cost_other_generation_regional_connection_costs)

