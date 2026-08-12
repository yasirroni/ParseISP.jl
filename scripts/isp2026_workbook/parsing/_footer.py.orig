# %% [markdown]
# ## Validation
#
# The checks below confirm worksheet coverage and the three source contracts already verified for ParseISP.

# %%
assert list(SHEET_RANGES) == workbook.sheetnames[:len(SHEET_RANGES)]
assert network_capability_flow_path_transfer_capability.shape == (18, 10) if len(SHEET_RANGES) >= 32 else True
assert transmission_reliability_transmission_unplanned_outage_rates.shape == (7, 4) if len(SHEET_RANGES) >= 34 else True
assert flow_path_augmentation_options_flow_path_augmentation_options.shape[0] == 62 if len(SHEET_RANGES) >= 38 else True

