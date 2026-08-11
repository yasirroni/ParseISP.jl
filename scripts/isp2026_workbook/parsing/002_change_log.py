# %% [markdown]
# ## Change Log
#
# Records changes made across workbook releases.

# %% [markdown]
# ### Workbook change log
#
# Records workbook versions, dates, changes, and supporting detail.
#
# Source block: `Change Log!B2:E657` (656 rows × 4 columns).

# %%
change_log_workbook_change_log = parse_spec('Change Log', {'name': 'Workbook change log', 'range': 'B2:E657'})
show_table(change_log_workbook_change_log)

