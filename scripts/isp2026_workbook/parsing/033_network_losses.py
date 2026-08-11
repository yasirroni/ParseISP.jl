# %% [markdown]
# ## Network losses
#
# Proportion of interconnector losses applied to regional reference nodes and loss equations.

# %% [markdown]
# ### Existing flow-path loss equations
#
# Source block: `Network losses!B5:J28` (24 rows × 9 columns).

# %%
network_losses_existing_flow_path_loss_equations = parse_spec('Network losses', {'name': 'Existing flow-path loss equations', 'range': 'B5:J28'})
show_table(network_losses_existing_flow_path_loss_equations)

# %% [markdown]
# ### Committed and anticipated project loss equations
#
# Source block: `Network losses!B30:J34` (5 rows × 9 columns).

# %%
network_losses_committed_and_anticipated_project_loss_equations = parse_spec('Network losses', {'name': 'Committed and anticipated project loss equations', 'range': 'B30:J34'})
show_table(network_losses_committed_and_anticipated_project_loss_equations)

# %% [markdown]
# ### Development-option loss equations
#
# Source block: `Network losses!B36:J88` (53 rows × 9 columns).

# %%
network_losses_development_option_loss_equations = parse_spec('Network losses', {'name': 'Development-option loss equations', 'range': 'B36:J88'})
show_table(network_losses_development_option_loss_equations)

