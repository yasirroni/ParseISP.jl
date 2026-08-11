# %% [markdown]
# ## Fuel Price Summary
#
# Summary (calculated) of generator fuel costs.

# %%
show_sheet_discovery('Fuel Price Summary')

# %% [markdown]
# ### Fuel-price scenario selection
#
# Maps the selected ISP scenario to the fuel-price summary calculations.
#
# Candidate source block: `B7:S9` (3 rows × 18 columns).

# %%
inspect_candidate('Fuel Price Summary', 'B7:S9')

# %% [markdown]
# ### Existing generator fuel prices
#
# Summarises fuel prices for existing, committed, anticipated, and additional generators.
#
# Candidate source block: `B11:AK738` (728 rows × 36 columns).

# %%
inspect_candidate('Fuel Price Summary', 'B11:AK738')

# %% [markdown]
# ### New entrant fuel prices
#
# Summarises fuel prices for new entrant generation technologies.
#
# Candidate source block: `B743:AK1268` (526 rows × 36 columns).

# %%
inspect_candidate('Fuel Price Summary', 'B743:AK1268')

