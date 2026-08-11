# %% [markdown]
# ## Assumptions Summary
#
# Provides workbook metadata, worksheet descriptions, and supporting materials.

# %% [markdown]
# ### Version history
#
# Records workbook version numbers, dates, and descriptions.
#
# Source block: `Assumptions Summary!B6:D35` (30 rows × 3 columns).

# %%
assumptions_summary_version_history = parse_spec('Assumptions Summary', {'name': 'Version history', 'range': 'B6:D35'})
show_table(assumptions_summary_version_history)

# %% [markdown]
# ### Worksheet descriptions
#
# Maps worksheets to assumption groups, descriptions, and sources.
#
# Source block: `Assumptions Summary!B37:E119` (83 rows × 4 columns).

# %%
assumptions_summary_worksheet_descriptions = parse_spec('Assumptions Summary', {'name': 'Worksheet descriptions', 'range': 'B37:E119'})
show_table(assumptions_summary_worksheet_descriptions)

# %% [markdown]
# ### Supporting materials
#
# Lists supporting source materials referenced by the workbook.
#
# Source block: `Assumptions Summary!B121:C147` (27 rows × 2 columns).

# %%
assumptions_summary_supporting_materials = parse_spec('Assumptions Summary', {'name': 'Supporting materials', 'range': 'B121:C147'})
show_table(assumptions_summary_supporting_materials)

