# %% [markdown]
# ## Max Ramp Rates
#
# Maximum rates of change for thermal unit output up and down.

# %% [markdown]
# ### Existing thermal generator maximum ramp rates
#
# Source block: `Max Ramp Rates!B7:F191` (185 rows × 5 columns).

# %%
max_ramp_rates_existing_thermal_generator_maximum_ramp_rates = parse_spec('Max Ramp Rates', {'name': 'Existing thermal generator maximum ramp rates', 'range': 'B7:F191'})
show_table(max_ramp_rates_existing_thermal_generator_maximum_ramp_rates)

# %% [markdown]
# ### New entrant maximum ramp rates
#
# Source block: `Max Ramp Rates!H7:J30` (24 rows × 3 columns).

# %%
max_ramp_rates_new_entrant_maximum_ramp_rates = parse_spec('Max Ramp Rates', {'name': 'New entrant maximum ramp rates', 'range': 'H7:J30'})
show_table(max_ramp_rates_new_entrant_maximum_ramp_rates)

