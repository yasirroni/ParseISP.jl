# %% [markdown]
# ## Fixed OPEX
#
# Fixed operating cost regardless of output for each generator or generator class.

# %% [markdown]
# ### Existing generator fixed OPEX
#
# Source block: `Fixed OPEX!B5:E739` (735 rows × 4 columns).

# %%
fixed_opex_existing_generator_fixed_opex = parse_spec('Fixed OPEX', {'name': 'Existing generator fixed OPEX', 'range': 'B5:E739'})
show_table(fixed_opex_existing_generator_fixed_opex)

# %% [markdown]
# ### New entrant fixed OPEX
#
# Source block: `Fixed OPEX!G5:I32` (28 rows × 3 columns).

# %%
fixed_opex_new_entrant_fixed_opex = parse_spec('Fixed OPEX', {'name': 'New entrant fixed OPEX', 'range': 'G5:I32'})
show_table(fixed_opex_new_entrant_fixed_opex)

