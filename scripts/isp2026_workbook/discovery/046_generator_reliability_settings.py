# %% [markdown]
# ## Generator Reliability Settings
#
# Defines a generators' unplanned outage rate, mean time to repair after an outage, and the derating experienced during a partial outage.

# %%
show_sheet_discovery('Generator Reliability Settings')

# %% [markdown]
# ### Existing generator long-duration outages
#
# Candidate source block: `B9:M16` (8 rows × 12 columns).

# %%
inspect_candidate('Generator Reliability Settings', 'B9:M16')

# %% [markdown]
# ### Existing generator outage rates and MTTR
#
# Candidate source block: `B21:M60` (40 rows × 12 columns).

# %%
inspect_candidate('Generator Reliability Settings', 'B21:M60')

# %% [markdown]
# ### New entrant reliability settings
#
# Candidate source block: `B62:H90` (29 rows × 7 columns).

# %%
inspect_candidate('Generator Reliability Settings', 'B62:H90')

