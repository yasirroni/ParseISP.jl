# %% [markdown]
# ## Economic Growth Forecasts
#
# Forecasts of Gross State Product (GSP) and Household Disposable Income (HDI).

# %% [markdown]
# ### Consultant forecast mapping
#
# Maps ISP scenarios to the consultant economic-growth scenarios.
#
# Source block: `Economic Growth Forecasts!B5:E8` (4 rows × 4 columns).

# %%
economic_growth_forecasts_consultant_forecast_mapping = parse_spec('Economic Growth Forecasts', {'name': 'Consultant forecast mapping', 'range': 'B5:E8'})
show_table(economic_growth_forecasts_consultant_forecast_mapping)

# %% [markdown]
# ### Gross State Product — Slower Growth
#
# Source block: `Economic Growth Forecasts!B12:AG19` (8 rows × 32 columns).

# %%
economic_growth_forecasts_gross_state_product_slower_growth = parse_spec('Economic Growth Forecasts', {'name': 'Gross State Product — Slower Growth', 'range': 'B12:AG19'})
show_table(economic_growth_forecasts_gross_state_product_slower_growth)

# %% [markdown]
# ### Gross State Product — Step Change
#
# Source block: `Economic Growth Forecasts!B21:AG28` (8 rows × 32 columns).

# %%
economic_growth_forecasts_gross_state_product_step_change = parse_spec('Economic Growth Forecasts', {'name': 'Gross State Product — Step Change', 'range': 'B21:AG28'})
show_table(economic_growth_forecasts_gross_state_product_step_change)

# %% [markdown]
# ### Gross State Product — Accelerated Transition
#
# Source block: `Economic Growth Forecasts!B30:AG37` (8 rows × 32 columns).

# %%
economic_growth_forecasts_gross_state_product_accelerated_transition = parse_spec('Economic Growth Forecasts', {'name': 'Gross State Product — Accelerated Transition', 'range': 'B30:AG37'})
show_table(economic_growth_forecasts_gross_state_product_accelerated_transition)

# %% [markdown]
# ### Household Disposable Income — Slower Growth
#
# Source block: `Economic Growth Forecasts!B41:AG48` (8 rows × 32 columns).

# %%
economic_growth_forecasts_household_disposable_income_slower_growth = parse_spec('Economic Growth Forecasts', {'name': 'Household Disposable Income — Slower Growth', 'range': 'B41:AG48'})
show_table(economic_growth_forecasts_household_disposable_income_slower_growth)

# %% [markdown]
# ### Household Disposable Income — Step Change
#
# Source block: `Economic Growth Forecasts!B50:AG57` (8 rows × 32 columns).

# %%
economic_growth_forecasts_household_disposable_income_step_change = parse_spec('Economic Growth Forecasts', {'name': 'Household Disposable Income — Step Change', 'range': 'B50:AG57'})
show_table(economic_growth_forecasts_household_disposable_income_step_change)

# %% [markdown]
# ### Household Disposable Income — Accelerated Transition
#
# Source block: `Economic Growth Forecasts!B59:AG66` (8 rows × 32 columns).

# %%
economic_growth_forecasts_household_disposable_income_accelerated_transition = parse_spec('Economic Growth Forecasts', {'name': 'Household Disposable Income — Accelerated Transition', 'range': 'B59:AG66'})
show_table(economic_growth_forecasts_household_disposable_income_accelerated_transition)

