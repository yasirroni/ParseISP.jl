# %% [markdown]
# ## Elec. Retail Price Indices
#
# Provides residential electricity retail price indices relative to the base year.

# %% [markdown]
# ### NEM residential electricity price index
#
# Provides residential retail price indices by ISP scenario.
#
# Source block: `Elec. Retail Price Indices!B8:AG12` (5 rows × 32 columns).

# %%
elec_retail_price_indices_nem_residential_electricity_price_index = parse_spec('Elec. Retail Price Indices', {'name': 'NEM residential electricity price index', 'range': 'B8:AG12'})
show_table(elec_retail_price_indices_nem_residential_electricity_price_index)

