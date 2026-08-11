# %% [markdown]
# ## Variable OPEX
#
# Variable operating cost per MWh of output for each generator or generator class.

# %% [markdown]
# ### Existing generator variable OPEX
#
# Source block: `Variable OPEX!B5:E738` (734 rows × 4 columns).

# %%
variable_opex_existing_generator_variable_opex = parse_spec('Variable OPEX', {'name': 'Existing generator variable OPEX', 'range': 'B5:E738'})
show_table(variable_opex_existing_generator_variable_opex)

# %% [markdown]
# ### New entrant variable OPEX
#
# Source block: `Variable OPEX!G5:H32` (28 rows × 2 columns).

# %%
variable_opex_new_entrant_variable_opex = parse_spec('Variable OPEX', {'name': 'New entrant variable OPEX', 'range': 'G5:H32'})
show_table(variable_opex_new_entrant_variable_opex)

