# %% [markdown]
# ## Carbon Budgets
#
# Global mean temperature increases by 2100 aligned with each scenario, and cumulative carbon budgets over the period to 2050.

# %%
show_sheet_discovery('Carbon Budgets')

# %% [markdown]
# ### NEM-wide carbon budgets
#
# Sets cumulative NEM-wide carbon budgets by scenario.
#
# Candidate source block: `B5:E9` (5 rows × 4 columns).

# %%
inspect_candidate('Carbon Budgets', 'B5:E9')

# %% [markdown]
# ### State carbon targets
#
# Records jurisdictional carbon targets used by the workbook.
#
# Candidate source block: `B15:G21` (7 rows × 6 columns).

# %%
inspect_candidate('Carbon Budgets', 'B15:G21')

# %% [markdown]
# ### Converted state carbon budgets
#
# Expresses state carbon targets in the workbook carbon-budget form.
#
# Candidate source block: `B25:D31` (7 rows × 3 columns).

# %%
inspect_candidate('Carbon Budgets', 'B25:D31')

