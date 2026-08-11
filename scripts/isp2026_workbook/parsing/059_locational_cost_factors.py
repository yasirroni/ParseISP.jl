# %% [markdown]
# ## Locational Cost Factors
#
# Locational cost factors provide an indication of the variation in new entrants generators cost based on the shift in labour, equipment and shipping/delivery cost between regions.

# %% [markdown]
# ### Non-pumped-hydro locational cost factors
#
# Source block: `Locational Cost Factors!B9:H80` (72 rows × 7 columns).

# %%
locational_cost_factors_non_pumped_hydro_locational_cost_factors = parse_spec('Locational Cost Factors', {'name': 'Non-pumped-hydro locational cost factors', 'range': 'B9:H80'})
show_table(locational_cost_factors_non_pumped_hydro_locational_cost_factors)

# %% [markdown]
# ### Pumped-hydro locational cost factors
#
# Source block: `Locational Cost Factors!B83:I132` (50 rows × 8 columns).

# %%
locational_cost_factors_pumped_hydro_locational_cost_factors = parse_spec('Locational Cost Factors', {'name': 'Pumped-hydro locational cost factors', 'range': 'B83:I132'})
show_table(locational_cost_factors_pumped_hydro_locational_cost_factors)

# %% [markdown]
# ### Technology cost breakdown ratios
#
# Source block: `Locational Cost Factors!B134:G158` (25 rows × 6 columns).

# %%
locational_cost_factors_technology_cost_breakdown_ratios = parse_spec('Locational Cost Factors', {'name': 'Technology cost breakdown ratios', 'range': 'B134:G158'})
show_table(locational_cost_factors_technology_cost_breakdown_ratios)

# %% [markdown]
# ### Technology-specific locational cost factors
#
# Source block: `Locational Cost Factors!B161:X227` (67 rows × 23 columns).

# %%
locational_cost_factors_technology_specific_locational_cost_factors = parse_spec('Locational Cost Factors', {'name': 'Technology-specific locational cost factors', 'range': 'B161:X227'})
show_table(locational_cost_factors_technology_specific_locational_cost_factors)

