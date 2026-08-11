# %% [markdown]
# ## Heat rates
#
# Efficiency of conversion of fuel to output for thermal generators.

# %% [markdown]
# ### Existing generator heat rates
#
# Source block: `Heat rates!B7:E740` (734 rows × 4 columns).

# %%
heat_rates_existing_generator_heat_rates = parse_spec('Heat rates', {'name': 'Existing generator heat rates', 'range': 'B7:E740'})
show_table(heat_rates_existing_generator_heat_rates)

# %% [markdown]
# ### New entrant heat rates
#
# Source block: `Heat rates!H7:I31` (25 rows × 2 columns).

# %%
heat_rates_new_entrant_heat_rates = parse_spec('Heat rates', {'name': 'New entrant heat rates', 'range': 'H7:I31'})
show_table(heat_rates_new_entrant_heat_rates)

