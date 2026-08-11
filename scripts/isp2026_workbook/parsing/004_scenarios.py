# %% [markdown]
# ## Scenarios
#
# Summary of scenario dimensions and parameters.

# %% [markdown]
# ### Scenario parameters
#
# Compares the main parameter settings across the three ISP scenarios.
#
# Source block: `Scenarios!B5:E29` (25 rows × 4 columns).

# %%
scenarios_scenario_parameters = parse_spec('Scenarios', {'name': 'Scenario parameters', 'range': 'B5:E29'})
show_table(scenarios_scenario_parameters)

