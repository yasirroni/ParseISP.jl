# %% [markdown]
# ## Maximum capacity
#
# Installed capacity of existing, committed and anticipated generators.

# %% [markdown]
# ### Existing, committed, anticipated, and additional generator capacity
#
# Source block: `Maximum capacity!B9:J750` (742 rows × 9 columns).

# %%
maximum_capacity_existing_committed_anticipated_and_additional_generator_capacity = parse_spec('Maximum capacity', {'name': 'Existing, committed, anticipated, and additional generator capacity', 'range': 'B9:J750'})
show_table(maximum_capacity_existing_committed_anticipated_and_additional_generator_capacity)

# %% [markdown]
# ### New generation technology capacity
#
# Source block: `Maximum capacity!L9:O31` (23 rows × 4 columns).

# %%
maximum_capacity_new_generation_technology_capacity = parse_spec('Maximum capacity', {'name': 'New generation technology capacity', 'range': 'L9:O31'})
show_table(maximum_capacity_new_generation_technology_capacity)

