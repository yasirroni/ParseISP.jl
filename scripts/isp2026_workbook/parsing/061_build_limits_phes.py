# %% [markdown]
# ## Build limits - PHES
#
# Modelled limitations for PHES impacting build constraints within the expansion modelling.

# %% [markdown]
# ### Pumped-hydro build limits
#
# Source block: `Build limits - PHES!B2:W27` (26 rows × 22 columns).

# %%
build_limits_phes_pumped_hydro_build_limits = parse_spec('Build limits - PHES', {'name': 'Pumped-hydro build limits', 'range': 'B2:W27'})
show_table(build_limits_phes_pumped_hydro_build_limits)

