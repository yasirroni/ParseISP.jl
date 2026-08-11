# %% [markdown]
# ## Carbon Budgets
#
# Global mean temperature increases by 2100 aligned with each scenario, and cumulative carbon budgets over the period to 2050.

# %% [markdown]
# ### NEM-wide carbon budgets
#
# Sets cumulative NEM-wide carbon budgets by scenario.
#
# Source block: `Carbon Budgets!B5:E9` (5 rows × 4 columns).

# %%
carbon_budgets_nem_wide_carbon_budgets = parse_spec('Carbon Budgets', {'name': 'NEM-wide carbon budgets', 'range': 'B5:E9'})
show_table(carbon_budgets_nem_wide_carbon_budgets)

# %% [markdown]
# ### State carbon targets
#
# Records jurisdictional carbon targets used by the workbook.
#
# Source block: `Carbon Budgets!B15:G21` (7 rows × 6 columns).

# %%
carbon_budgets_state_carbon_targets = parse_spec('Carbon Budgets', {'name': 'State carbon targets', 'range': 'B15:G21'})
show_table(carbon_budgets_state_carbon_targets)

# %% [markdown]
# ### Converted state carbon budgets
#
# Expresses state carbon targets in the workbook carbon-budget form.
#
# Source block: `Carbon Budgets!B25:D31` (7 rows × 3 columns).

# %%
carbon_budgets_converted_state_carbon_budgets = parse_spec('Carbon Budgets', {'name': 'Converted state carbon budgets', 'range': 'B25:D31'})
show_table(carbon_budgets_converted_state_carbon_budgets)

