# %% [markdown]
# ## Coverage check
#
# The catalogue must account for every worksheet in workbook order. A worksheet can legitimately have no
# embedded semantic data table; those cases remain explicit rather than being silently omitted.

# %%
assert list(SHEET_RANGES) == workbook_formula.sheetnames[:len(SHEET_RANGES)]
sum(len(ranges) for ranges in SHEET_RANGES.values())

