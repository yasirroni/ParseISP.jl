# %% [markdown]
# ## Retirement
#
# Announced and end-of-technical-life generator retirement.

# %% [markdown]
# ### Expected generator closure years
#
# Source block: `Retirement!B8:F738` (731 rows × 5 columns).

# %%
retirement_expected_generator_closure_years = parse_spec('Retirement', {'name': 'Expected generator closure years', 'range': 'B8:F738'})
show_table(retirement_expected_generator_closure_years)

# %% [markdown]
# ### Generator retirement costs
#
# Source block: `Retirement!H8:I50` (43 rows × 2 columns).

# %%
retirement_generator_retirement_costs = parse_spec('Retirement', {'name': 'Generator retirement costs', 'range': 'H8:I50'})
show_table(retirement_generator_retirement_costs)

