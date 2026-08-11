# %% [markdown]
# ## Gas, Liquid fuel, H2 price
#
# Fuel price for each gas and liquid fuel generator.

# %% [markdown]
# ### Existing GPG fuel costs
#
# Source block: `Gas, Liquid fuel, H2 price!B7:AG129` (123 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_existing_gpg_fuel_costs = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Existing GPG fuel costs', 'range': 'B7:AG129'})
show_table(gas_liquid_fuel_h2_price_existing_gpg_fuel_costs)

# %% [markdown]
# ### New entrant GPG fuel costs
#
# Source block: `Gas, Liquid fuel, H2 price!B132:AG224` (93 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_new_entrant_gpg_fuel_costs = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'New entrant GPG fuel costs', 'range': 'B132:AG224'})
show_table(gas_liquid_fuel_h2_price_new_entrant_gpg_fuel_costs)

# %% [markdown]
# ### Industrial fuel costs
#
# Source block: `Gas, Liquid fuel, H2 price!B228:AG249` (22 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_industrial_fuel_costs = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Industrial fuel costs', 'range': 'B228:AG249'})
show_table(gas_liquid_fuel_h2_price_industrial_fuel_costs)

# %% [markdown]
# ### Residential and commercial fuel costs
#
# Source block: `Gas, Liquid fuel, H2 price!B253:AG274` (22 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_residential_and_commercial_fuel_costs = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Residential and commercial fuel costs', 'range': 'B253:AG274'})
show_table(gas_liquid_fuel_h2_price_residential_and_commercial_fuel_costs)

# %% [markdown]
# ### Liquid fuel prices
#
# Source block: `Gas, Liquid fuel, H2 price!B278:AG302` (25 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_liquid_fuel_prices = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Liquid fuel prices', 'range': 'B278:AG302'})
show_table(gas_liquid_fuel_h2_price_liquid_fuel_prices)

# %% [markdown]
# ### GPG secondary liquid-fuel prices
#
# Source block: `Gas, Liquid fuel, H2 price!B305:AG429` (125 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_gpg_secondary_liquid_fuel_prices = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'GPG secondary liquid-fuel prices', 'range': 'B305:AG429'})
show_table(gas_liquid_fuel_h2_price_gpg_secondary_liquid_fuel_prices)

# %% [markdown]
# ### Hydrogen prices
#
# Source block: `Gas, Liquid fuel, H2 price!B433:AG438` (6 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_hydrogen_prices = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Hydrogen prices', 'range': 'B433:AG438'})
show_table(gas_liquid_fuel_h2_price_hydrogen_prices)

# %% [markdown]
# ### Biomethane prices
#
# Source block: `Gas, Liquid fuel, H2 price!B440:AG452` (13 rows × 32 columns).

# %%
gas_liquid_fuel_h2_price_biomethane_prices = parse_spec('Gas, Liquid fuel, H2 price', {'name': 'Biomethane prices', 'range': 'B440:AG452'})
show_table(gas_liquid_fuel_h2_price_biomethane_prices)

