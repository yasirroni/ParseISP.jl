# %% [markdown]
# ## Financial parameters
#
# Financial parameters (discount rate, weighted average cost of capital, value of customer reliability, and value of emissions reductions) used during cost benefit analysis.

# %% [markdown]
# ### Discount rate
#
# Source block: `Financial parameters!B2:F7` (6 rows × 5 columns).

# %%
financial_parameters_discount_rate = parse_spec('Financial parameters', {'name': 'Discount rate', 'range': 'B2:F7'})
show_table(financial_parameters_discount_rate)

# %% [markdown]
# ### Weighted Average Cost of Capital
#
# Source block: `Financial parameters!B10:F41` (32 rows × 5 columns).

# %%
financial_parameters_weighted_average_cost_of_capital = parse_spec('Financial parameters', {'name': 'Weighted Average Cost of Capital', 'range': 'B10:F41'})
show_table(financial_parameters_weighted_average_cost_of_capital)

# %% [markdown]
# ### Value of Customer Reliability
#
# Source block: `Financial parameters!B43:G51` (9 rows × 6 columns).

# %%
financial_parameters_value_of_customer_reliability = parse_spec('Financial parameters', {'name': 'Value of Customer Reliability', 'range': 'B43:G51'})
show_table(financial_parameters_value_of_customer_reliability)

# %% [markdown]
# ### Value of emissions reduction
#
# Source block: `Financial parameters!B54:C90` (37 rows × 2 columns).

# %%
financial_parameters_value_of_emissions_reduction = parse_spec('Financial parameters', {'name': 'Value of emissions reduction', 'range': 'B54:C90'})
show_table(financial_parameters_value_of_emissions_reduction)

