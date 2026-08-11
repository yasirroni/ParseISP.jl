# %% [markdown]
# ## Hydro Scheme Inflows
#
# Monthly aggregated inflow trends for reference years.

# %% [markdown]
# ### Secondary hydro scheme releases and outflows
#
# Source block: `Hydro Scheme Inflows!B4:T79` (76 rows × 19 columns).

# %%
hydro_scheme_inflows_secondary_hydro_scheme_releases_and_outflows = parse_spec('Hydro Scheme Inflows', {'name': 'Secondary hydro scheme releases and outflows', 'range': 'B4:T79'})
show_table(hydro_scheme_inflows_secondary_hydro_scheme_releases_and_outflows)

# %% [markdown]
# ### Run-of-river hydro outflows
#
# Source block: `Hydro Scheme Inflows!B81:T121` (41 rows × 19 columns).

# %%
hydro_scheme_inflows_run_of_river_hydro_outflows = parse_spec('Hydro Scheme Inflows', {'name': 'Run-of-river hydro outflows', 'range': 'B81:T121'})
show_table(hydro_scheme_inflows_run_of_river_hydro_outflows)

# %% [markdown]
# ### Hydro Tasmania scheme
#
# Source block: `Hydro Scheme Inflows!B123:T141` (19 rows × 19 columns).

# %%
hydro_scheme_inflows_hydro_tasmania_scheme = parse_spec('Hydro Scheme Inflows', {'name': 'Hydro Tasmania scheme', 'range': 'B123:T141'})
show_table(hydro_scheme_inflows_hydro_tasmania_scheme)

# %% [markdown]
# ### Snowy Hydro weather-variability representation
#
# Source block: `Hydro Scheme Inflows!B143:T162` (20 rows × 19 columns).

# %%
hydro_scheme_inflows_snowy_hydro_weather_variability_representation = parse_spec('Hydro Scheme Inflows', {'name': 'Snowy Hydro weather-variability representation', 'range': 'B143:T162'})
show_table(hydro_scheme_inflows_snowy_hydro_weather_variability_representation)

