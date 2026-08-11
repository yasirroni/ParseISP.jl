# %% [markdown]
# ## Power System Security
#
# Reflect power system constraints to reflect secure operating limits.

# %% [markdown]
# ### Coal-retirement minimum-fault-level costs
#
# Source block: `Power System Security!B4:D49` (46 rows × 3 columns).

# %%
power_system_security_coal_retirement_minimum_fault_level_costs = parse_spec('Power System Security', {'name': 'Coal-retirement minimum-fault-level costs', 'range': 'B4:D49'})
show_table(power_system_security_coal_retirement_minimum_fault_level_costs)

# %% [markdown]
# ### Efficient system-strength costs
#
# Source block: `Power System Security!B52:AE56` (5 rows × 30 columns).

# %%
power_system_security_efficient_system_strength_costs = parse_spec('Power System Security', {'name': 'Efficient system-strength costs', 'range': 'B52:AE56'})
show_table(power_system_security_efficient_system_strength_costs)

# %% [markdown]
# ### Synchronous unit commitment — standard scenarios
#
# Source block: `Power System Security!B58:G72` (15 rows × 6 columns).

# %%
power_system_security_synchronous_unit_commitment_standard_scenarios = parse_spec('Power System Security', {'name': 'Synchronous unit commitment — standard scenarios', 'range': 'B58:G72'})
show_table(power_system_security_synchronous_unit_commitment_standard_scenarios)

# %% [markdown]
# ### Synchronous unit commitment — Accelerated Transition
#
# Source block: `Power System Security!B74:G94` (21 rows × 6 columns).

# %%
power_system_security_synchronous_unit_commitment_accelerated_transition = parse_spec('Power System Security', {'name': 'Synchronous unit commitment — Accelerated Transition', 'range': 'B74:G94'})
show_table(power_system_security_synchronous_unit_commitment_accelerated_transition)

