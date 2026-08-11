# %% [markdown]
# ## Maintenance
#
# The percentage of time per year that a generator is expected to be out of service for maintenance. De-rating applied to generators under maintenance, staged construction or to track age-related degradation.

# %% [markdown]
# ### Existing generator maintenance rates
#
# Source block: `Maintenance!B5:D29` (25 rows × 3 columns).

# %%
maintenance_existing_generator_maintenance_rates = parse_spec('Maintenance', {'name': 'Existing generator maintenance rates', 'range': 'B5:D29'})
show_table(maintenance_existing_generator_maintenance_rates)

# %% [markdown]
# ### New entrant maintenance rates
#
# Source block: `Maintenance!G5:I32` (28 rows × 3 columns).

# %%
maintenance_new_entrant_maintenance_rates = parse_spec('Maintenance', {'name': 'New entrant maintenance rates', 'range': 'G5:I32'})
show_table(maintenance_new_entrant_maintenance_rates)

