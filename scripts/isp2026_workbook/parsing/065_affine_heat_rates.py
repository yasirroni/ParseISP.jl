# %% [markdown]
# ## Affine Heat rates
#
# Heat rate curves for large thermal units.

# %% [markdown]
# ### Existing generator affine heat rates
#
# Source block: `Affine Heat rates!B6:F192` (187 rows × 5 columns).

# %%
affine_heat_rates_existing_generator_affine_heat_rates = parse_spec('Affine Heat rates', {'name': 'Existing generator affine heat rates', 'range': 'B6:F192'})
show_table(affine_heat_rates_existing_generator_affine_heat_rates)

# %% [markdown]
# ### New entrant affine heat rates
#
# Source block: `Affine Heat rates!H6:K29` (24 rows × 4 columns).

# %%
affine_heat_rates_new_entrant_affine_heat_rates = parse_spec('Affine Heat rates', {'name': 'New entrant affine heat rates', 'range': 'H6:K29'})
show_table(affine_heat_rates_new_entrant_affine_heat_rates)

