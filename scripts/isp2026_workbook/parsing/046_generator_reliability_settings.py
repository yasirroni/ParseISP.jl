# %% [markdown]
# ## Generator Reliability Settings
#
# Defines a generators' unplanned outage rate, mean time to repair after an outage, and the derating experienced during a partial outage.

# %% [markdown]
# ### Existing generator long-duration outages
#
# Source block: `Generator Reliability Settings!B9:M16` (8 rows × 12 columns).

# %%
generator_reliability_settings_existing_generator_long_duration_outages = parse_spec('Generator Reliability Settings', {'name': 'Existing generator long-duration outages', 'range': 'B9:M16'})
show_table(generator_reliability_settings_existing_generator_long_duration_outages)

# %% [markdown]
# ### Existing generator outage rates and MTTR
#
# Source block: `Generator Reliability Settings!B21:M60` (40 rows × 12 columns).

# %%
generator_reliability_settings_existing_generator_outage_rates_and_mttr = parse_spec('Generator Reliability Settings', {'name': 'Existing generator outage rates and MTTR', 'range': 'B21:M60'})
show_table(generator_reliability_settings_existing_generator_outage_rates_and_mttr)

# %% [markdown]
# ### New entrant reliability settings
#
# Source block: `Generator Reliability Settings!B62:H90` (29 rows × 7 columns).

# %%
generator_reliability_settings_new_entrant_reliability_settings = parse_spec('Generator Reliability Settings', {'name': 'New entrant reliability settings', 'range': 'B62:H90'})
show_table(generator_reliability_settings_new_entrant_reliability_settings)

