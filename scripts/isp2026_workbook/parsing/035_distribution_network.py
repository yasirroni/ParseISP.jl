# %% [markdown]
# ## Distribution network
#
# Inputs used to model distribution network opportunities to facilitate aggregate operation of consumer energy resources and other distributed resources.

# %% [markdown]
# ### Mid-scale generation and storage build limits
#
# Source block: `Distribution network!B11:G38` (28 rows × 6 columns).

# %%
distribution_network_mid_scale_generation_and_storage_build_limits = parse_spec('Distribution network', {'name': 'Mid-scale generation and storage build limits', 'range': 'B11:G38'})
show_table(distribution_network_mid_scale_generation_and_storage_build_limits)

# %% [markdown]
# ### Distribution CER augmentation tranche costs
#
# Source block: `Distribution network!B40:H57` (18 rows × 7 columns).

# %%
distribution_network_distribution_cer_augmentation_tranche_costs = parse_spec('Distribution network', {'name': 'Distribution CER augmentation tranche costs', 'range': 'B40:H57'})
show_table(distribution_network_distribution_cer_augmentation_tranche_costs)

# %% [markdown]
# ### Average CER generation-limit time-of-day profile
#
# Source block: `Distribution network!B59:AZ1433` (1375 rows × 51 columns).

# %%
distribution_network_average_cer_generation_limit_time_of_day_profile = parse_spec('Distribution network', {'name': 'Average CER generation-limit time-of-day profile', 'range': 'B59:AZ1433'})
show_table(distribution_network_average_cer_generation_limit_time_of_day_profile)

