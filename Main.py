#%%
import pandas as pd
import numpy as np
import statsmodels.formula.api as smf
import matplotlib.pyplot as plt
import seaborn as sns
import os
import datetime as dt
import scipy.stats as sp
from matplotlib.pyplot import savefig
import openpyxl
from statsmodels.graphics.tukeyplot import results
import matplotlib as mpl
import folium
import io
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from matplotlib.path import Path
from matplotlib.patches import PathPatch
from mpl_toolkits.axes_grid1.axes_divider import make_axes_locatable
import matplotlib.ticker as mtick
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import fonts
import plotly
import plotly.graph_objects as go
import plotly.express as px
import kaleido
import dash
import micropip
from dash.exceptions import PreventUpdate
from dash import Dash, html, dcc, Input, Output, callback, no_update, State
from dash.dependencies import Input, Output
import dash_ag_grid as dag # imports data to display in a grid table
import dash_bootstrap_components as dbc
import dash_mantine_components as dmc
from pandas.plotting import scatter_matrix
# from matplotlib import pyplot as plotly
from matplotlib import font_manager as font_manager
from datetime import datetime,timedelta
import plotly.figure_factory as ff
os.chdir("C:\\Mike's Stuff\\Hockey Data\\Five-Year Data")
#%%
              ### First Steps for App

#1 - import data (A3Z, Evolving Hockey, Nat Stat)
#2 - clean data and create columns - groupby team, player by season --- create per60 columns
#%%
              ### Read In 5v5 Data

alltrack = pd.read_excel('Tracking2126.xlsx', sheet_name = 'CleanedColumns')
onice = pd.read_excel('5v5OnIce.xlsx', sheet_name = '5v5OI')
prod = pd.read_excel('5v5Prod.xlsx', sheet_name = 'Prod')
rel = pd.read_excel('5v5RelTm.xlsx', sheet_name = 'RelTm')
#%%
              ### Replace spaces with _ and replace 'ARI' with 'UTA' and replace 'C/L/R' with 'F'

alltrack.columns = alltrack.columns.str.replace(' ', '_')
alltrack['Team'] = alltrack['Team'].replace({'ARI' : 'UTA'})
alltrack['Team1'] = alltrack['Team1'].replace({'ARI' : 'UTA'})
alltrack['Team2'] = alltrack['Team2'].replace({'ARI' : 'UTA'})
alltrack['Position'] = alltrack['Pos.'].replace({'C' : 'F', 'L' : 'F', 'R' : 'F', 'l' : 'F', 'f' : 'F'})
#%%
              ### Groupby Team, Season, Position and Create New Columns

teamtrack = alltrack.groupby(['Team', 'Season', 'Position'])[['Chances', 'Primary_Shot_Assists', '5v5_TOI',
       'Chance_Assists', 'Home_Plate', 'Low_to_High', 'Behind_Net',
       'Center_Lane_Assists', 'NZ_Assist', 'DZ_Assist', 'Shots_off_Rush',
       'Assists_off_Rush', 'Shots_off_Forecheck', 'Assists_off_Forecheck',
       'Shots_off_Cycle', 'Assists_off_Cycle', 'Shots_off_HD_Passes',
       'Zone_Entries', 'Carries', 'Failed_Entries', 'Entries_w/_Passing_Play',
       'Recoveries', 'Carries_w_Chances', 'Dump_in_Chances',
       'Forecheck_Pressures', 'DZ_Puck_Touches', 'DZ_Retrievals', 'Zone_Exits',
       'Exits_w_Possession', 'Carried_Exits', 'Passed_Exits', 'Clears',
       'Missed_Passes', 'Retrievals_Leading_to_Exits', 'Botched_Retrievals',
       'Exchanges', 'Failed_Exit', 'Rushed_Exits', 'Second_Touch_Exits',
       'Targets', 'Carries.1', 'Denials', 'Passes_Allowed',
       'Carries_w_Chance_Against', 'Dump-in_w_Chance_Against']].sum(numeric_only = True).reset_index()

teamtrack['HD_Passes'] = teamtrack['Home_Plate'] + teamtrack['Behind_Net']
teamtrack['Chance_Contributions'] = teamtrack['Chances'] + teamtrack['Chance_Assists']
teamtrack['Rush_Contributions'] = teamtrack['Shots_off_Rush'] + teamtrack['Assists_off_Rush']
teamtrack['Cycle_Fcheck_Contributions'] = teamtrack['Shots_off_Forecheck'] + teamtrack['Assists_off_Forecheck'] + teamtrack['Shots_off_Cycle'] + teamtrack['Assists_off_Cycle']
teamtrack['Carry_Entry_Pct'] = teamtrack['Carries'] / teamtrack['Zone_Entries']
teamtrack['Botch_Fail_Pct'] = (teamtrack['Botched_Retrievals']  / teamtrack['DZ_Retrievals']) + (teamtrack['Failed_Exit'] / teamtrack['Zone_Exits'])
#%%
teamtrack.columns
#%%
              ### Create Pivot Table to Sort by Team, Position, and Season

teampivot = pd.pivot_table(teamtrack, index = ['Team', 'Season', 'Position'],
                           values = ['5v5_TOI', 'Chances', 'Chance_Assists', 'HD_Passes', 'Chance_Contributions', 'Rush_Contributions',
                                     'Cycle_Fcheck_Contributions', 'Zone_Entries', 'Carries', 'DZ_Retrievals', 'Zone_Exits', 'Botched_Retrievals', 'Failed_Exit']).reset_index()
#%%
              ### Function to Create Per-60 Rates by Team, Position, and Season

def team60(teampivot):
    stat_cols = ['Botched_Retrievals', 'Carries', 'Chance_Assists', 'Chance_Contributions', 'Chances', 'Cycle_Fcheck_Contributions', 'DZ_Retrievals', 'Failed_Exit', 'HD_Passes', 'Rush_Contributions',
                 'Zone_Entries', 'Zone_Exits']
    grouped = (teampivot.groupby(['Team', 'Season', 'Position'], as_index = False).sum())
    for col in stat_cols:
        grouped[f'{col}_Per_60'] = (grouped[col] / grouped['5v5_TOI']) * 60
    return grouped
grouped = team60(teampivot)
print(grouped)
#%%
                     ### Add Carry_Entry_Pct and Botch_Fail_Pct

grouped['Botch_Fail_Pct'] = (grouped['Botched_Retrievals'] / grouped['DZ_Retrievals'] + grouped['Failed_Exit'] / grouped['Zone_Exits'])
grouped['Carry_Entry_Pct'] = grouped['Carries'] / grouped['Zone_Entries']
#%%
                     ### Create Function for Relative to Lg Avg

def calculate_league_60_rates(df):

    stat_cols = ['Botched_Retrievals', 'Carries', 'Chance_Assists', 'Chance_Contributions', 'Chances', 'Cycle_Fcheck_Contributions', 'DZ_Retrievals', 'Failed_Exit', 'HD_Passes', 'Rush_Contributions',
                        'Zone_Entries', 'Zone_Exits']

    ### Group by Season and Position, Then Calcualte Per-60 League Averages
    league_rates = (
        df.groupby(['Season', 'Position'])
          .apply(
              lambda g: pd.Series({
                  f'{stat}_60_league': (g[stat].sum() / g['5v5_TOI'].sum()) * 60
                  for stat in stat_cols
              })
          )
    )

    return league_rates


def compare_team_to_league(df):
       ### Compare Team Stats to League Average

    stat_cols = ['Botched_Retrievals', 'Carries', 'Chance_Assists', 'Chance_Contributions', 'Chances', 'Cycle_Fcheck_Contributions', 'DZ_Retrievals', 'Failed_Exit', 'HD_Passes', 'Rush_Contributions',
                        'Zone_Entries', 'Zone_Exits']

    # Team-level per-60 rates
    team_rates = (
        df.groupby(['Season', 'Position', 'Team'])
          .apply(
              lambda g: pd.Series({
                  f'{stat}_60_team': (g[stat].sum() / g['5v5_TOI'].sum()) * 60
                  for stat in stat_cols
              })
          )
          .reset_index()
    )

    # Create League Averages Using First Function
    league_rates = (
        calculate_league_60_rates(df)
        .reset_index()
    )

    # Merge Team Per-60 with League Averages
    result = team_rates.merge(
        league_rates,
        on=['Season', 'Position'],
        how='left'
    )

    # Create Relative Metrics
    for stat in stat_cols:
        result[f'{stat}_relative'] = (
            result[f'{stat}_60_team'] -
            result[f'{stat}_60_league']
        ) / result[f'{stat}_60_league']

    return result
team_compare = compare_team_to_league(teampivot)
print(team_compare)
#%%
team_compare.columns
#%%
team_compare = team_compare.reset_index()
#%%
team_compare.head()
#%%
team_compare['Season'].astype
#%%
            ### Melt 'team_compare'
# team_melt = pd.melt(team_compare, id_vars = ['Team'], var_name = 'category', value_name = 'relative', value_vars =
#                     ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                      'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'])
# season_melt = pd.melt(team_compare, id_vars = ['Season'], var_name = 'campaigns', value_name = 'years')
#%%
            ### Compile App - Attempt 1

# fig = px.line_polar(team_melt, r = 'relative', theta = 'category', color = 'Team', line_close = True, render_mode = 'svg',
#                     hover_name = 'Team', hover_data = {'Team' : False}, markers = True, direction = 'clockwise', start_angle = 180)
# #
# #
# # # Initialize App
# app = Dash(__name__)
# #
# # # App Layout
# app.layout = html.Div([
#     html.Div(children = 'Team Relative Tracking Data By Season, 2021-2026'),
#     html.Hr(),
#     dcc.Dropdown(options = team_melt.Team.unique(), placeholder = "Select Team", multi = True, id = 'teams', value = "ANA"),
#     dcc.Dropdown(options = season_melt.Season.unique(), id = 'seasons', value = "2021-22"),
#     dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                             'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
#                  placeholder = "Choose Stat", multi = True, id = 'stats', value = "Relative Scoring Chances"),
#     dbc.Button('Submit', id = 'my_button', n_clicks = 0),
#     dcc.Graph(figure = fig, id = "graph1")
# ])
#
# @callback(
#     Output(component_id = "graph1", component_property = "figure"),
#     Input(component_id = "my_button", component_property = "n_clicks"),
#     State(component_id = "teams", component_property = "placeholder"),
#     State(component_id = "seasons", component_property = "placeholder"),
#     State(component_id = "stats", component_property = "placeholder"),
#     prevent_initial_call = True
# )
#
# def update_layout(teams, season, stats):
#     if teams is None:
#         return no_update
#     if season is None:
#         return no_update
#     if stats is None:
#         return no_update
#
# if __name__ == '__main__':
#     app.run(debug=True)

#%%
            ### Melt 'team_compare'
team_melt2 = pd.melt(team_compare, id_vars = ['Team', 'Season'], var_name = 'category', value_name = 'relative', value_vars =
                    ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
                     'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'])
#%%
            ### Compile App - Attempt 2

# fig = px.line_polar(team_melt, r = 'relative', theta = 'category', color = 'Team', line_close = True, render_mode = 'svg',
#                     hover_name = 'Team', hover_data = {'Team' : False}, markers = True, direction = 'clockwise', start_angle = 180)
#
#
# # Initialize App
app = Dash(__name__)
#
# # App Layout
app.layout = html.Div([
    html.Div(children = 'Team Relative Tracking Data By Season, 2021-2026', style={"color" : "white"}),
    html.Hr(),
    dcc.Dropdown(options = team_melt2.Team.unique(), placeholder = "Select Team", multi = True, id = 'teams', value = ["ANA"]),
    dcc.Dropdown(options = team_melt2.Season.unique(), id = 'seasons', placeholder = "Select Season", value = "2021-22"),
    dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
                            'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
                 placeholder = "Choose Stat", multi = True, id = 'stats', value = "Relative Chances"),
    dbc.Button('Submit', id = 'my_button', n_clicks = 0),
    dcc.Graph(id = "graph1")
])

### Add Controls for interaction

@callback(
    Output(component_id = "graph1", component_property = "figure"),
    Input("my_button", "n_clicks"),
    State(component_id = "teams", component_property = "value"),
    State(component_id = "seasons", component_property = "value"),
    State(component_id = "stats", component_property = "value"),
    prevent_initial_call = True
)

def update_graph(button, teams, season, stats):
    filtered_df = team_melt2[
        (team_melt2['Team'].isin(teams)) &
        (team_melt2['Season'] == season)]
    fig = px.line_polar(filtered_df, r = 'relative', theta = 'category', color = 'Team', line_close = True,
                        hover_name = 'Team', markers = True, direction = 'clockwise', start_angle = 180)
    return fig

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)
#%%
