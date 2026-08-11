# %% [markdown]
# ## Build limits - REZs
#
# Modelled limitations for REZs impacting build constraints within the expansion modelling.

# %% [markdown]
# ### Initial REZ resource limits
#
# Source block: `Build limits - REZs!B2:Q62` (61 rows × 16 columns).

# %%
build_limits_rezs_initial_rez_resource_limits = parse_spec('Build limits - REZs', {'name': 'Initial REZ resource limits', 'range': 'B2:Q62'})
show_table(build_limits_rezs_initial_rez_resource_limits)

# %% [markdown]
# ### Initial REZ transmission limits
#
# Source block: `Build limits - REZs!B64:N119` (56 rows × 13 columns).

# %%
build_limits_rezs_initial_rez_transmission_limits = parse_spec('Build limits - REZs', {'name': 'Initial REZ transmission limits', 'range': 'B64:N119'})
show_table(build_limits_rezs_initial_rez_transmission_limits)

# %% [markdown]
# ### REZ transmission modifiers
#
# Source block: `Build limits - REZs!B121:F132` (12 rows × 5 columns).

# %%
build_limits_rezs_rez_transmission_modifiers = parse_spec('Build limits - REZs', {'name': 'REZ transmission modifiers', 'range': 'B121:F132'})
show_table(build_limits_rezs_rez_transmission_modifiers)

# %% [markdown]
# ### REZ group constraints
#
# Source block: `Build limits - REZs!B136:K265` (130 rows × 10 columns).

# %%
build_limits_rezs_rez_group_constraints = parse_spec('Build limits - REZs', {'name': 'REZ group constraints', 'range': 'B136:K265'})
show_table(build_limits_rezs_rez_group_constraints)

# %% [markdown]
# ### REZ transmission limit constraints
#
# Source block: `Build limits - REZs!B267:K317` (51 rows × 10 columns).

# %%
build_limits_rezs_rez_transmission_limit_constraints = parse_spec('Build limits - REZs', {'name': 'REZ transmission limit constraints', 'range': 'B267:K317'})
show_table(build_limits_rezs_rez_transmission_limit_constraints)

# %% [markdown]
# ### REZ secondary transmission limits
#
# Source block: `Build limits - REZs!B319:K335` (17 rows × 10 columns).

# %%
build_limits_rezs_rez_secondary_transmission_limits = parse_spec('Build limits - REZs', {'name': 'REZ secondary transmission limits', 'range': 'B319:K335'})
show_table(build_limits_rezs_rez_secondary_transmission_limits)

# %% [markdown]
# ### Non-REZ connections pipeline build limits
#
# Source block: `Build limits - REZs!B337:E356` (20 rows × 4 columns).

# %%
build_limits_rezs_non_rez_connections_pipeline_build_limits = parse_spec('Build limits - REZs', {'name': 'Non-REZ connections pipeline build limits', 'range': 'B337:E356'})
show_table(build_limits_rezs_non_rez_connections_pipeline_build_limits)

# %% [markdown]
# ### REZ technology-specific access-right limits
#
# Source block: `Build limits - REZs!B358:G368` (11 rows × 6 columns).

# %%
build_limits_rezs_rez_technology_specific_access_right_limits = parse_spec('Build limits - REZs', {'name': 'REZ technology-specific access-right limits', 'range': 'B358:G368'})
show_table(build_limits_rezs_rez_technology_specific_access_right_limits)

