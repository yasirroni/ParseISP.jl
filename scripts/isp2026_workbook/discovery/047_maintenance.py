# %% [markdown]
# ## Maintenance
#
# The percentage of time per year that a generator is expected to be out of service for maintenance. De-rating applied to generators under maintenance, staged construction or to track age-related degradation.

# %%
show_sheet_discovery('Maintenance')

# %% [markdown]
# ### Existing generator maintenance rates
#
# Candidate source block: `B5:D29` (25 rows × 3 columns).

# %%
inspect_candidate('Maintenance', 'B5:D29')

# %% [markdown]
# ### New entrant maintenance rates
#
# Candidate source block: `G5:I32` (28 rows × 3 columns).

# %%
inspect_candidate('Maintenance', 'G5:I32')

