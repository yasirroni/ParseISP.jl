# %% [markdown]
# ## Regional Build Costs Summary
#
# Summary (calculated) of regional build costs for a selectable scenario / technology.

# %% [markdown]
# ### Build-cost selection
#
# Records the scenario and technology controls used by the regional build-cost summary.
#
# Source block: `Regional Build Costs Summary!B7:C10` (4 rows × 2 columns).

# %%
regional_build_costs_summary_build_cost_selection = parse_spec('Regional Build Costs Summary', {'name': 'Build-cost selection', 'range': 'B7:C10'})
show_table(regional_build_costs_summary_build_cost_selection)

# %% [markdown]
# ### Regional build costs
#
# Summarises regional build costs after locational cost factors are applied.
#
# Source block: `Regional Build Costs Summary!B12:AV75` (64 rows × 47 columns).

# %%
regional_build_costs_summary_regional_build_costs = parse_spec('Regional Build Costs Summary', {'name': 'Regional build costs', 'range': 'B12:AV75'})
show_table(regional_build_costs_summary_regional_build_costs)

