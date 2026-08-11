# %% [markdown]
# ## Auxiliary
#
# Auxiliary (self) load for each generator or generator class.

# %% [markdown]
# ### Existing generator auxiliary load
#
# Source block: `Auxiliary!B5:E736` (732 rows × 4 columns).

# %%
auxiliary_existing_generator_auxiliary_load = parse_spec('Auxiliary', {'name': 'Existing generator auxiliary load', 'range': 'B5:E736'})
show_table(auxiliary_existing_generator_auxiliary_load)

# %% [markdown]
# ### New entrant auxiliary load
#
# Source block: `Auxiliary!G5:H29` (25 rows × 2 columns).

# %%
auxiliary_new_entrant_auxiliary_load = parse_spec('Auxiliary', {'name': 'New entrant auxiliary load', 'range': 'G5:H29'})
show_table(auxiliary_new_entrant_auxiliary_load)

