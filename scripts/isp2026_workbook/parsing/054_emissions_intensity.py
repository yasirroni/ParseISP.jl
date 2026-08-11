# %% [markdown]
# ## Emissions intensity
#
# Emissions production per MWh of output for each generator or generator class.

# %% [markdown]
# ### Existing generator emissions intensity
#
# Source block: `Emissions intensity!B4:E744` (741 rows × 4 columns).

# %%
emissions_intensity_existing_generator_emissions_intensity = parse_spec('Emissions intensity', {'name': 'Existing generator emissions intensity', 'range': 'B4:E744'})
show_table(emissions_intensity_existing_generator_emissions_intensity)

# %% [markdown]
# ### New entrant emissions intensity
#
# Source block: `Emissions intensity!G4:H29` (26 rows × 2 columns).

# %%
emissions_intensity_new_entrant_emissions_intensity = parse_spec('Emissions intensity', {'name': 'New entrant emissions intensity', 'range': 'G4:H29'})
show_table(emissions_intensity_new_entrant_emissions_intensity)

