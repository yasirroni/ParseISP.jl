# %% [markdown]
# ## Hydrogen consumption locations
#
# Location of hydrogen consumption and candidate hydrogen hubs and ports.

# %% [markdown]
# ### Regional hydrogen consumption allocation rule
#
# Defines the priority rule used to locate regional hydrogen consumption for green commodities and export.
#
# Source block: `Hydrogen consumption locations!B7:C9` (3 rows × 2 columns).

# %%
hydrogen_consumption_locations_regional_allocation_rule = parse_spec('Hydrogen consumption locations', {'name': 'Regional hydrogen consumption allocation rule', 'range': 'B7:C9'})
show_table(hydrogen_consumption_locations_regional_allocation_rule)

# %% [markdown]
# ### Subregional hydrogen consumption allocation rule
#
# Defines the priority rule used to locate domestic hydrogen consumption within subregions.
#
# Source block: `Hydrogen consumption locations!B15:C18` (4 rows × 2 columns).

# %%
hydrogen_consumption_locations_subregional_allocation_rule = parse_spec('Hydrogen consumption locations', {'name': 'Subregional hydrogen consumption allocation rule', 'range': 'B15:C18'})
show_table(hydrogen_consumption_locations_subregional_allocation_rule)

# %% [markdown]
# ### Hydrogen consumption allocation
#
# Applies the allocation rules to green-commodity, export, and domestic hydrogen consumption by ISP subregion.
#
# Source block: `Hydrogen consumption locations!B24:F40` (17 rows × 5 columns).

# %%
hydrogen_consumption_locations_hydrogen_consumption_allocation = parse_spec('Hydrogen consumption locations', {'name': 'Hydrogen consumption allocation', 'range': 'B24:F40'})
show_table(hydrogen_consumption_locations_hydrogen_consumption_allocation)

# %% [markdown]
# ### Hydrogen hubs
#
# Source block: `Hydrogen consumption locations!B42:B44` (3 rows × 1 columns).

# %%
hydrogen_consumption_locations_hydrogen_hubs = parse_spec('Hydrogen consumption locations', {'name': 'Hydrogen hubs', 'range': 'B42:B44'})
show_table(hydrogen_consumption_locations_hydrogen_hubs)

# %% [markdown]
# ### Hydrogen and commodity export ports
#
# Source block: `Hydrogen consumption locations!B46:D57` (12 rows × 3 columns).

# %%
hydrogen_consumption_locations_hydrogen_and_commodity_export_ports = parse_spec('Hydrogen consumption locations', {'name': 'Hydrogen and commodity export ports', 'range': 'B46:D57'})
show_table(hydrogen_consumption_locations_hydrogen_and_commodity_export_ports)

