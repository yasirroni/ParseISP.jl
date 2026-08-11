# %% [markdown]
# ## ONSG
#
# Other non-scheduled generation (ONSG) capacity forecast.

# %% [markdown]
# ### Sub-regional ONSG capacity
#
# Source block: `ONSG!B8:AH55` (48 rows × 33 columns).

# %%
onsg_sub_regional_onsg_capacity = parse_spec('ONSG', {'name': 'Sub-regional ONSG capacity', 'range': 'B8:AH55'})
show_table(onsg_sub_regional_onsg_capacity)

# %% [markdown]
# ### Regional ONSG capacity
#
# Source block: `ONSG!B57:AG74` (18 rows × 32 columns).

# %%
onsg_regional_onsg_capacity = parse_spec('ONSG', {'name': 'Regional ONSG capacity', 'range': 'B57:AG74'})
show_table(onsg_regional_onsg_capacity)

