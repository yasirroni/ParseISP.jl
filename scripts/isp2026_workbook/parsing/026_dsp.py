# %% [markdown]
# ## DSP
#
# Demand side participation forecast.

# %% [markdown]
# ### Summer demand-side participation
#
# Source block: `DSP!B7:AI84` (78 rows × 34 columns).

# %%
dsp_summer_demand_side_participation = parse_spec('DSP', {'name': 'Summer demand-side participation', 'range': 'B7:AI84'})
show_table(dsp_summer_demand_side_participation)

# %% [markdown]
# ### Winter demand-side participation
#
# Source block: `DSP!B87:AI164` (78 rows × 34 columns).

# %%
dsp_winter_demand_side_participation = parse_spec('DSP', {'name': 'Winter demand-side participation', 'range': 'B87:AI164'})
show_table(dsp_winter_demand_side_participation)

