# %% [markdown]
# ## Existing Gen Data Summary
#
# Summary (calculated) of the generator technical data.

# %% [markdown]
# ### Existing generation data summary
#
# Summarises technical data for existing, committed, anticipated, and additional generators.
#
# Source block: `Existing Gen Data Summary!B10:AT738` (729 rows × 45 columns).

# %%
existing_gen_data_summary_existing_generation_data_summary = parse_spec('Existing Gen Data Summary', {'name': 'Existing generation data summary', 'range': 'B10:AT738'})
show_table(existing_gen_data_summary_existing_generation_data_summary)

