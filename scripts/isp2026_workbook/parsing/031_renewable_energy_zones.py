# %% [markdown]
# ## Renewable energy zones
#
# Renewable energy zones.

# %% [markdown]
# ### Candidate renewable energy zones
#
# Lists candidate REZ identifiers, names, NEM regions, and ISP sub-regions.
#
# Source block: `Renewable energy zones!B6:E53` (48 rows × 4 columns).

# %%
renewable_energy_zones_candidate_renewable_energy_zones = parse_spec('Renewable energy zones', {'name': 'Candidate renewable energy zones', 'range': 'B6:E53'})
show_table(renewable_energy_zones_candidate_renewable_energy_zones)

