# %% [markdown]
# ## Hydrogen demand - Domestic
#
# Projections of domestic hydrogen demand.

# %% [markdown]
# ### Domestic hydrogen demand
#
# Source block: `Hydrogen demand - Domestic!B2:AH53` (52 rows × 33 columns).

# %%
hydrogen_demand_domestic_domestic_hydrogen_demand = parse_spec('Hydrogen demand - Domestic', {'name': 'Domestic hydrogen demand', 'range': 'B2:AH53'})
show_table(hydrogen_demand_domestic_domestic_hydrogen_demand)

