# %% [markdown]
# ## Reserves
#
# Minimum reserve levels for reliable regional supply.

# %% [markdown]
# ### Initial regional reserves
#
# Source block: `Reserves!B2:C14` (13 rows × 2 columns).

# %%
reserves_initial_regional_reserves = parse_spec('Reserves', {'name': 'Initial regional reserves', 'range': 'B2:C14'})
show_table(reserves_initial_regional_reserves)

