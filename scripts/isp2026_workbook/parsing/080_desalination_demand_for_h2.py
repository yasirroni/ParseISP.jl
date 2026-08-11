# %% [markdown]
# ## Desalination demand for H2
#
# Projections of electricity required for water treatment associated with electrolytic hydrogen production.

# %% [markdown]
# ### Desalination electricity demand for hydrogen
#
# Source block: `Desalination demand for H2!B2:AH52` (51 rows × 33 columns).

# %%
desalination_demand_for_h2_desalination_electricity_demand_for_hydrogen = parse_spec('Desalination demand for H2', {'name': 'Desalination electricity demand for hydrogen', 'range': 'B2:AH52'})
show_table(desalination_demand_for_h2_desalination_electricity_demand_for_hydrogen)

