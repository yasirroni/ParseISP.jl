# %% [markdown]
# ## GPG Min Stable Level
#
# Minimum operating levels for large GPG units.

# %% [markdown]
# ### Existing GPG minimum stable levels
#
# Source block: `GPG Min Stable Level!B10:E150` (141 rows × 4 columns).

# %%
gpg_min_stable_level_existing_gpg_minimum_stable_levels = parse_spec('GPG Min Stable Level', {'name': 'Existing GPG minimum stable levels', 'range': 'B10:E150'})
show_table(gpg_min_stable_level_existing_gpg_minimum_stable_levels)

# %% [markdown]
# ### New entrant GPG minimum stable levels
#
# Source block: `GPG Min Stable Level!G10:H35` (26 rows × 2 columns).

# %%
gpg_min_stable_level_new_entrant_gpg_minimum_stable_levels = parse_spec('GPG Min Stable Level', {'name': 'New entrant GPG minimum stable levels', 'range': 'G10:H35'})
show_table(gpg_min_stable_level_new_entrant_gpg_minimum_stable_levels)

