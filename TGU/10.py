from pandas import *
from scipy import stats
import seaborn as sb
import matplotlib.pyplot as plt



df = read_csv("ab_test_results_aggregated_views_clicks_2.csv")
groups = df["group"].unique()

for group in groups:
    df_group = df[df["group"] == group]
    n = df_group["user_id"].count() #всего пользователей
    n_views = df_group["views"].sum()
    n_clicks = df_group["clicks"].sum()
    n_null_click = df_group[(df_group["views"] > 0) & (df_group["clicks"] == 0)]["user_id"].count()
    n_null_click = n_null_click / n_views * 100
    min_click = df_group["clicks"].min()
    max_click = df_group["clicks"].max()
    max_view = df_group["views"].max()
    print(f"{group:>8} views = {n_views:>10} clicks = {n_clicks:>10} null_click = {n_null_click:6.2f}% "
          f"max_view={max_view} min_click = {min_click} max_click = {max_click}")
