# %% [markdown]
# ## Electrification
#
# Electrification in all sectors excluding road transportation.

# %% [markdown]
# ### Consultant forecast mapping
#
# Source block: `Electrification!B7:E10` (4 rows × 4 columns).

# %%
electrification_consultant_forecast_mapping = parse_spec('Electrification', {'name': 'Consultant forecast mapping', 'range': 'B7:E10'})
show_table(electrification_consultant_forecast_mapping)

# %% [markdown]
# ### Electrification — Slower Growth
#
# Source block: `Electrification!B12:AF19` (8 rows × 31 columns).

# %%
electrification_electrification_slower_growth = parse_spec('Electrification', {'name': 'Electrification — Slower Growth', 'range': 'B12:AF19'})
show_table(electrification_electrification_slower_growth)

# %% [markdown]
# ### Electrification — Step Change
#
# Source block: `Electrification!B21:AF28` (8 rows × 31 columns).

# %%
electrification_electrification_step_change = parse_spec('Electrification', {'name': 'Electrification — Step Change', 'range': 'B21:AF28'})
show_table(electrification_electrification_step_change)

# %% [markdown]
# ### Electrification — Accelerated Transition
#
# Source block: `Electrification!B30:AF37` (8 rows × 31 columns).

# %%
electrification_electrification_accelerated_transition = parse_spec('Electrification', {'name': 'Electrification — Accelerated Transition', 'range': 'B30:AF37'})
show_table(electrification_electrification_accelerated_transition)

