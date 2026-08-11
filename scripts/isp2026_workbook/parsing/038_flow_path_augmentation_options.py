# %% [markdown]
# ## Flow path augmentation options
#
# Capability, cost and timing for flow path augmentation options.

# %% [markdown]
# ### Flow-path augmentation options
#
# Combines repeated physical flow-path sections into one logical option dataset.
#
# Source block: `Flow path augmentation options!B11:Q127` (117 rows × 16 columns).

# %%
flow_path_augmentation_options_flow_path_augmentation_options = parse_spec('Flow path augmentation options', {'name': 'Flow-path augmentation options',
 'range': 'B11:Q127',
 'parser': 'flow_path_options',
 'expected_semantic_rows': 62})
show_table(flow_path_augmentation_options_flow_path_augmentation_options)

