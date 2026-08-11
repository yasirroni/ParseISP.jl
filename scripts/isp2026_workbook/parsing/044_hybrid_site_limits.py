# %% [markdown]
# ## Hybrid site limits
#
# Interval-level charging/dispatch limits for sites with a combination of VRE and battery storage at one connection point.

# %% [markdown]
# ### Hybrid-site limits
#
# Contains the verified hybrid-site limit source table.
#
# Source block: `Hybrid site limits!B9:G67` (59 rows × 6 columns).

# %%
hybrid_site_limits_hybrid_site_limits = parse_spec('Hybrid site limits', {'name': 'Hybrid-site limits', 'range': 'B9:G67'})
show_table(hybrid_site_limits_hybrid_site_limits)

