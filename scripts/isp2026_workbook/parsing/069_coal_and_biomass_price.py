# %% [markdown]
# ## Coal and Biomass price
#
# Coal fuel price for each coal generator.

# %% [markdown]
# ### Coal fuel prices
#
# Source block: `Coal and Biomass price!B8:AG54` (47 rows × 32 columns).

# %%
coal_and_biomass_price_coal_fuel_prices = parse_spec('Coal and Biomass price', {'name': 'Coal fuel prices', 'range': 'B8:AG54'})
show_table(coal_and_biomass_price_coal_fuel_prices)

# %% [markdown]
# ### Biomass fuel prices
#
# Source block: `Coal and Biomass price!B57:AG61` (5 rows × 32 columns).

# %%
coal_and_biomass_price_biomass_fuel_prices = parse_spec('Coal and Biomass price', {'name': 'Biomass fuel prices', 'range': 'B57:AG61'})
show_table(coal_and_biomass_price_biomass_fuel_prices)

