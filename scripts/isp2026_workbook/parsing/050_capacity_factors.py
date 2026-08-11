# %% [markdown]
# ## Capacity Factors
#
# Capacity factors for renewable generators in the renewable energy zones.

# %% [markdown]
# ### New large-scale renewable capacity factors
#
# Source block: `Capacity Factors !B2:V214` (213 rows × 21 columns).

# %%
capacity_factors_new_large_scale_renewable_capacity_factors = parse_spec('Capacity Factors ', {'name': 'New large-scale renewable capacity factors', 'range': 'B2:V214'})
show_table(capacity_factors_new_large_scale_renewable_capacity_factors)

