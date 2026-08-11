# %% [markdown]
# ## Summary Mapping
#
# Master look-up table for generator, storage, and electrolyser assets.

# %% [markdown]
# ### Existing, committed, and anticipated asset mapping
#
# Source block: `Summary Mapping!C2:AF733` (732 rows × 30 columns).

# %%
summary_mapping_existing_committed_and_anticipated_asset_mapping = parse_spec('Summary Mapping', {'name': 'Existing, committed, and anticipated asset mapping', 'range': 'C2:AF733'})
show_table(summary_mapping_existing_committed_and_anticipated_asset_mapping)

# %% [markdown]
# ### Consumer energy resource mapping
#
# Source block: `Summary Mapping!C734:AF786` (53 rows × 30 columns).

# %%
summary_mapping_consumer_energy_resource_mapping = parse_spec('Summary Mapping', {'name': 'Consumer energy resource mapping', 'range': 'C734:AF786'})
show_table(summary_mapping_consumer_energy_resource_mapping)

# %% [markdown]
# ### New entrant asset mapping
#
# Source block: `Summary Mapping!C790:AF1316` (527 rows × 30 columns).

# %%
summary_mapping_new_entrant_asset_mapping = parse_spec('Summary Mapping', {'name': 'New entrant asset mapping', 'range': 'C790:AF1316'})
show_table(summary_mapping_new_entrant_asset_mapping)

# %% [markdown]
# ### New entrant electrolyser mapping
#
# Source block: `Summary Mapping!C1319:AF1381` (63 rows × 30 columns).

# %%
summary_mapping_new_entrant_electrolyser_mapping = parse_spec('Summary Mapping', {'name': 'New entrant electrolyser mapping', 'range': 'C1319:AF1381'})
show_table(summary_mapping_new_entrant_electrolyser_mapping)

