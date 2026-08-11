# %% [markdown]
# ## End use fuel consumption data
#
# Data for end-use fuel consumption by scenario across the NEM chart, identified by multi-sectoral modelling conducted by CSIRO (Figure 1 in IASR).

# %% [markdown]
# ### End-use fuel consumption — Slower Growth
#
# Source block: `End use fuel consumption data!B6:AF15` (10 rows × 31 columns).

# %%
end_use_fuel_consumption_data_end_use_fuel_consumption_slower_growth = parse_spec('End use fuel consumption data', {'name': 'End-use fuel consumption — Slower Growth', 'range': 'B6:AF15'})
show_table(end_use_fuel_consumption_data_end_use_fuel_consumption_slower_growth)

# %% [markdown]
# ### End-use fuel consumption — Step Change
#
# Source block: `End use fuel consumption data!B17:AF26` (10 rows × 31 columns).

# %%
end_use_fuel_consumption_data_end_use_fuel_consumption_step_change = parse_spec('End use fuel consumption data', {'name': 'End-use fuel consumption — Step Change', 'range': 'B17:AF26'})
show_table(end_use_fuel_consumption_data_end_use_fuel_consumption_step_change)

# %% [markdown]
# ### End-use fuel consumption — Accelerated Transition
#
# Source block: `End use fuel consumption data!B28:AF37` (10 rows × 31 columns).

# %%
end_use_fuel_consumption_data_end_use_fuel_consumption_accelerated_transition = parse_spec('End use fuel consumption data', {'name': 'End-use fuel consumption — Accelerated Transition', 'range': 'B28:AF37'})
show_table(end_use_fuel_consumption_data_end_use_fuel_consumption_accelerated_transition)

