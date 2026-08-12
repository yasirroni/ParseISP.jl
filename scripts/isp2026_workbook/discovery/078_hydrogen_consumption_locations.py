# %% [markdown]
# ## Hydrogen consumption locations
#
# Location of hydrogen consumption and candidate hydrogen hubs and ports.

# %%
show_sheet_discovery('Hydrogen consumption locations')

# %% [markdown]
# ### Regional hydrogen consumption allocation rule
#
# Defines the priority rule used to locate regional hydrogen consumption for green commodities and export.
#
# Source block: `Hydrogen consumption locations!B7:C9` (3 rows × 2 columns).

# %%
inspect_source_range('Hydrogen consumption locations', 'B7:C9')

# %% [markdown]
# ### Subregional hydrogen consumption allocation rule
#
# Defines the priority rule used to locate domestic hydrogen consumption within subregions.
#
# Source block: `Hydrogen consumption locations!B15:C18` (4 rows × 2 columns).

# %%
inspect_source_range('Hydrogen consumption locations', 'B15:C18')

# %% [markdown]
# ### Hydrogen consumption allocation
#
# Applies the allocation rules to green-commodity, export, and domestic hydrogen consumption by ISP subregion.
#
# Source block: `Hydrogen consumption locations!B24:F40` (17 rows × 5 columns).

# %%
inspect_source_range('Hydrogen consumption locations', 'B24:F40')

# %% [markdown]
# ### Hydrogen hubs
#
# Candidate source block: `B42:B44` (3 rows × 1 columns).

# %%
inspect_candidate('Hydrogen consumption locations', 'B42:B44')

# %% [markdown]
# ### Hydrogen and commodity export ports
#
# Candidate source block: `B46:D57` (12 rows × 3 columns).

# %%
inspect_candidate('Hydrogen consumption locations', 'B46:D57')

