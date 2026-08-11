# %% [markdown]
# ## Fuel Price Summary
#
# Summary (calculated) of generator fuel costs.

# %% [markdown]
# ### Fuel-price scenario selection
#
# Maps the selected ISP scenario to the fuel-price summary calculations.
#
# Source block: `Fuel Price Summary!B7:S9` (3 rows × 18 columns).

# %%
fuel_price_summary_fuel_price_scenario_selection = parse_spec('Fuel Price Summary', {'name': 'Fuel-price scenario selection', 'range': 'B7:S9'})
show_table(fuel_price_summary_fuel_price_scenario_selection)

# %% [markdown]
# ### Existing generator fuel prices
#
# Summarises fuel prices for existing, committed, anticipated, and additional generators.
#
# Source block: `Fuel Price Summary!B11:AK738` (728 rows × 36 columns).

# %%
fuel_price_summary_existing_generator_fuel_prices = parse_spec('Fuel Price Summary', {'name': 'Existing generator fuel prices', 'range': 'B11:AK738'})
show_table(fuel_price_summary_existing_generator_fuel_prices)

# %% [markdown]
# ### New entrant fuel prices
#
# Summarises fuel prices for new entrant generation technologies.
#
# Source block: `Fuel Price Summary!B743:AK1268` (526 rows × 36 columns).

# %%
fuel_price_summary_new_entrant_fuel_prices = parse_spec('Fuel Price Summary', {'name': 'New entrant fuel prices', 'range': 'B743:AK1268'})
show_table(fuel_price_summary_new_entrant_fuel_prices)

