# %% [markdown]
# ## Seasonal ratings
#
# Winter and summer ratings for existing, committed, advanced and new entrant generators.

# %% [markdown]
# ### New generation technology seasonal ratings
#
# Source block: `Seasonal ratings!B9:E36` (28 rows × 4 columns).

# %%
seasonal_ratings_new_generation_technology_seasonal_ratings = parse_spec('Seasonal ratings', {'name': 'New generation technology seasonal ratings', 'range': 'B9:E36'})
show_table(seasonal_ratings_new_generation_technology_seasonal_ratings)

# %% [markdown]
# ### Existing, committed, anticipated, and additional generator seasonal ratings
#
# Source block: `Seasonal ratings!B42:AI770` (729 rows × 34 columns).

# %%
seasonal_ratings_existing_committed_anticipated_and_additional_generator_seasonal_ratings = parse_spec('Seasonal ratings', {'name': 'Existing, committed, anticipated, and additional generator seasonal ratings', 'range': 'B42:AI770'})
show_table(seasonal_ratings_existing_committed_anticipated_and_additional_generator_seasonal_ratings)

