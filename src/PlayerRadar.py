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
              ### Read In 5v5 Data

alltrack = pd.read_excel('Tracking2126.xlsx', sheet_name = 'CleanedColumns')
#%%
              ### Replace spaces with _ and replace 'ARI' with 'UTA' and replace 'C/L/R' with 'F'

alltrack.columns = alltrack.columns.str.replace(' ', '_')
alltrack['Team'] = alltrack['Team'].replace({'ARI' : 'UTA'})
alltrack['Team1'] = alltrack['Team1'].replace({'ARI' : 'UTA'})
alltrack['Team2'] = alltrack['Team2'].replace({'ARI' : 'UTA'})
alltrack['Position'] = alltrack['Pos.'].replace({'C' : 'F', 'L' : 'F', 'R' : 'F', 'l' : 'F', 'f' : 'F'})
#%%
              ### Groupby Player, Team, Season, Position and Create New Columns

playertrack = alltrack.groupby(['Player', 'Team', 'Season', 'Position'])[['Chances', 'Primary_Shot_Assists', '5v5_TOI',
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

playertrack['HD_Passes'] = playertrack['Home_Plate'] + playertrack['Behind_Net']
playertrack['Chance_Contributions'] = playertrack['Chances'] + playertrack['Chance_Assists']
playertrack['Rush_Contributions'] = playertrack['Shots_off_Rush'] + playertrack['Assists_off_Rush']
playertrack['Cycle_Fcheck_Contributions'] = playertrack['Shots_off_Forecheck'] + playertrack['Assists_off_Forecheck'] + playertrack['Shots_off_Cycle'] + playertrack['Assists_off_Cycle']
playertrack['Carry_Entry_Pct'] = playertrack['Carries'] / playertrack['Zone_Entries']
playertrack['Botch_Fail_Pct'] = (playertrack['Botched_Retrievals']  / playertrack['DZ_Retrievals']) + (playertrack['Failed_Exit'] / playertrack['Zone_Exits'])
#%%
              ### Create Pivot Table to Sort by Team, Position, and Season

playerpivot = pd.pivot_table(playertrack, index = ['Team', 'Player', 'Season', 'Position'],
                           values = ['5v5_TOI', 'Chances', 'Chance_Assists', 'HD_Passes', 'Chance_Contributions', 'Rush_Contributions',
                                     'Cycle_Fcheck_Contributions', 'Zone_Entries', 'Carries', 'DZ_Retrievals', 'Zone_Exits', 'Botched_Retrievals', 'Failed_Exit']).reset_index()
#%%
              ### Function to Create Per-60 Rates by Team, Position, and Season

def player60(playerpivot):
    stat_cols = ['Botched_Retrievals', 'Carries', 'Chance_Assists', 'Chance_Contributions', 'Chances', 'Cycle_Fcheck_Contributions', 'DZ_Retrievals', 'Failed_Exit', 'HD_Passes', 'Rush_Contributions',
                 'Zone_Entries', 'Zone_Exits']
    grouped = (playerpivot.groupby(['Team', 'Player', 'Season', 'Position'], as_index = False).sum())
    for col in stat_cols:
        grouped[f'{col}_Per_60'] = (grouped[col] / grouped['5v5_TOI']) * 60
    return grouped
grouped = player60(playerpivot)
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


def compare_player_to_league(df):
       ### Compare Team Stats to League Average

    stat_cols = ['Botched_Retrievals', 'Carries', 'Chance_Assists', 'Chance_Contributions', 'Chances', 'Cycle_Fcheck_Contributions', 'DZ_Retrievals', 'Failed_Exit', 'HD_Passes', 'Rush_Contributions',
                        'Zone_Entries', 'Zone_Exits']

    # Player-level Per-60 Rates
    player_rates = (
        df.groupby(['Season', 'Position', 'Player'])
          .apply(
              lambda g: pd.Series({
                  f'{stat}_60_player': (g[stat].sum() / g['5v5_TOI'].sum()) * 60
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
    result = player_rates.merge(
        league_rates,
        on=['Season', 'Position'],
        how='left'
    )

    # Create Relative Metrics
    for stat in stat_cols:
        result[f'{stat}_relative'] = (
            result[f'{stat}_60_player'] -
            result[f'{stat}_60_league']
        ) / result[f'{stat}_60_league']

    return result
player_compare = compare_player_to_league(playerpivot)
player_compare = player_compare.reset_index()
#%%
            ### Melt 'player_compare'
player_melt = pd.melt(player_compare, id_vars=['Season', 'Player', 'Position'], var_name='category', value_name='relative',
                    value_vars=
                    ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
                     'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'])
#%%
#             ### Combine Player Relative App - WORKS
#
# # fig = px.line_polar(team_melt, r = 'relative', theta = 'category', color = 'Team', line_close = True, render_mode = 'svg',
# #                     hover_name = 'Team', hover_data = {'Team' : False}, markers = True, direction = 'clockwise', start_angle = 180)
#
#
# ### Initialize App
# app = Dash(__name__)
# server = app.server
#
# ### App Layout
# app.layout = html.Div([
#     html.Div(children = 'Player Relative Tracking Data By Season, 2021-2026', style={"color" : "Black"}),
#     html.Hr(),
#     dcc.Dropdown(options = player_melt.Player.unique(), id = 'players', placeholder = "Select Player", multi = True, value = ["5ebastian Aho"]),
#     dcc.Dropdown(options = player_melt.Season.unique(), id = 'seasons', placeholder = "Select Season", value = "2021-22"),
#     dcc.Dropdown(options = player_melt.Position.unique(), id = 'positions', placeholder = "Select Position", multi = True, value = player_melt.Position.unique().tolist()),
#     dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                             'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
#                  placeholder = "Choose Stat", multi = True, id = 'stats', value = "Relative Chances"),
#     html.Div(id = 'stats-error', style = {"color" : "blue"}),
#     dbc.Button('Submit', id = 'my_button', n_clicks = 0),
#     dcc.Graph(id = "graph1")
# ])
#
# ### Add Controls for interaction
#
# @callback(
#     Output(component_id = "graph1", component_property = "figure"),
#     Output(component_id = "stats-error", component_property = "children"),
#     Input("my_button", "n_clicks"),
#     State(component_id = "players", component_property = "value"),
#     State(component_id = "seasons", component_property = "value"),
#     State(component_id = "positions", component_property = "value"),
#     State(component_id = "stats", component_property = "value"),
#     prevent_initial_call = True
# )
#
# def update_graph(button, players, season, positions, stats):
#     if stats is None or len(stats) < 3:
#         return {}, 'Please Choose At Least 3 Stats'
#
#     filtered_df = player_melt[
#         (player_melt['Player'].isin(players)) &
#         (player_melt['Season'] == season) &
#         (player_melt['Position'].isin(positions)) &
#         (player_melt['category'].isin(stats))]
#     fig = px.line_polar(filtered_df, r = 'relative', theta = 'category', color = 'Player', line_close = True,
#                         hover_name = 'Player', markers = True, direction = 'clockwise', start_angle = 180, render_mode = 'svg')
#     return fig, ""
#
# if __name__ == '__main__':
#     app.run(debug=True, use_reloader=False)
#%%
#             ### Combine Player Relative App with Position Callback - WORKS
#
# app = Dash(__name__)
# server = app.server
#
# ### App Layout
# app.layout = html.Div([
#     html.Div(children = 'Player Relative Tracking Data By Season, 2021-2026', style={"color" : "Black"}),
#     html.Hr(),
#     dcc.Dropdown(id = 'players', placeholder = "Select Player", multi = True),
#     dcc.Dropdown(options = player_melt.Season.unique(), id = 'seasons', placeholder = "Select Season", value = "2021-22"),
#     dcc.RadioItems(options = [{"label" : pos, "value" : pos}
#                               for pos in player_melt.Position.unique()], id = 'positions', value = 'forward', inline = True),
#     dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                             'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
#                  placeholder = "Choose Stat", multi = True, id = 'stats', value = "Relative Chances"),
#     html.Div(id = 'stats-error', style = {"color" : "blue"}),
#     dbc.Button('Submit', id = 'my_button', n_clicks = 0),
#     dcc.Graph(id = "graph1")
# ])
#
# ### Update Player Dropdown Menus Based On Position
#
# @callback(
#     Output('players', 'options'),
#     Output('players', 'value'),
#     Input('positions', 'value')
# )
#
# def update_player(select_pos):
#     filtered_players = player_melt[player_melt['Position'] == select_pos]['Player'].unique()
#     options = [{'label': player, 'value': player} for player in filtered_players]
#     value = [filtered_players[0]] if len(filtered_players) > 0 else []
#     return options, value
# ### Add Controls for interaction
#
# @callback(
#     Output(component_id = "graph1", component_property = "figure"),
#     Output(component_id = "stats-error", component_property = "children"),
#     Input("my_button", "n_clicks"),
#     State(component_id = "players", component_property = "value"),
#     State(component_id = "seasons", component_property = "value"),
#     State(component_id = "positions", component_property = "value"),
#     State(component_id = "stats", component_property = "value"),
#     prevent_initial_call = True
# )
#
# def update_graph(button, players, season, position, stats):
#     if stats is None or len(stats) < 3:
#         return {}, 'Please Choose At Least 3 Stats'
#
#     filtered_df = player_melt[
#         (player_melt['Player'].isin(players)) &
#         (player_melt['Season'] == season) &
#         (player_melt['Position'] == position) &
#         (player_melt['category'].isin(stats))]
#     fig = px.line_polar(filtered_df, r = 'relative', theta = 'category', color = 'Player', line_close = True,
#                         hover_name = 'Player', markers = True, direction = 'clockwise', start_angle = 180, render_mode = 'svg')
#     return fig, ""
#
# if __name__ == '__main__':
#     app.run(debug=True, use_reloader=False)
#%%
#             ### Same Relative Player App but Moving Radio Button Location with Radio Button Callback - WORKS
#
# app = Dash(__name__)
# server = app.server
#
# ### App Layout
# app.layout = html.Div([
#     html.Div(children = 'Player Relative Tracking Data By Season, 2021-2026', style={"color" : "Black"}),
#     html.Hr(),
#     html.H4(children = 'Choose Position', style={"color" : "Black"}),
#     dcc.RadioItems(options = [{"label" : pos, "value" : pos}
#                               for pos in player_melt.Position.unique()], id = 'positions', value = 'forward', inline = True),
#     html.Hr(),
#     dcc.Dropdown(id = 'players', placeholder = "Select Player", multi = True, value = ['Sidney Crosby']),
#     dcc.Dropdown(options = player_melt.Season.unique(), id = 'seasons', placeholder = "Select Season", value = "2021-22"),
#     dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                             'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
#                  placeholder = "Choose Stat", multi = True, id = 'stats',
#                  value = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
#                           'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative']),
#     html.Div(id = 'stats-error', style = {"color" : "blue"}),
#     dbc.Button('Submit', id = 'my_button', n_clicks = 0),
#     dcc.Graph(id = "graph1")
# ])
#
# ### Update Player Dropdown Menus Based On Position
#
# @callback(
#     Output('players', 'options'),
#     Output('players', 'value'),
#     Input('positions', 'value'),
# )
#
# def update_player(select_pos):
#     filtered_players = (player_melt[player_melt['Position'] == select_pos]['Player'].unique())
#     options = [{'label': player, 'value': player} for player in filtered_players]
#     value = [filtered_players[0]] if len(filtered_players) > 0 else []
#     return options, value
# ### Add Controls for interaction
#
# @callback(
#     Output(component_id = "graph1", component_property = "figure"),
#     Output(component_id = "stats-error", component_property = "children"),
#     Input("my_button", "n_clicks"),
#     State(component_id = "players", component_property = "value"),
#     State(component_id = "seasons", component_property = "value"),
#     State(component_id = "positions", component_property = "value"),
#     State(component_id = "stats", component_property = "value"),
#     prevent_initial_call = True
# )
#
# def update_graph(button, players, season, position, stats):
#     if stats is None or len(stats) < 3:
#         return {}, 'Please Choose At Least 3 Stats'
#
#     filtered_df = player_melt[
#         (player_melt['Player'].isin(players)) &
#         (player_melt['Season'] == season) &
#         (player_melt['Position'] == position) &
#         (player_melt['category'].isin(stats))]
#     fig = px.line_polar(filtered_df, r = 'relative', theta = 'category', color = 'Player', line_close = True,
#                         hover_name = 'Player', markers = True, direction = 'clockwise', start_angle = 180, render_mode = 'svg')
#     return fig, ""
#
# if __name__ == '__main__':
#     app.run(debug=True, use_reloader=False)
#%%
            ### Same Relative Player App With New Radio Button Location with Radio Button Callback - WORKS
            ### Adding A Season Chained Callback to Update Player List

app = Dash(__name__)
server = app.server

### App Layout
app.layout = html.Div([
    html.Div(children = 'Player Relative Tracking Data By Season, 2021-2026', style={"color" : "Black"}),
    html.Hr(),
    html.H4(children = 'Choose Position', style={"color" : "Black"}),
    dcc.RadioItems(options = [{"label" : pos, "value" : pos}
                              for pos in player_melt.Position.unique()], id = 'positions', value = 'forward', inline = True),
    html.Hr(),
    html.H4(children = 'Select Season', style={"color" : "Black"}),
    dcc.Dropdown(options = player_melt.Season.unique(), id = 'seasons', placeholder = "Select Season", value = "2021-22"),
    html.Hr(),
    html.H4(children = 'Select Player(s)', style={"color" : "Black"}),
    dcc.Dropdown(id = 'players', placeholder = "Select Player", multi = True, value = ['Sidney Crosby']),
    html.Hr(),
    html.H4(children = 'Select Stats', style={"color" : "Black"}),
    dcc.Dropdown(options = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
                            'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative'],
                 placeholder = "Choose Stat", multi = True, id = 'stats',
                 value = ['Chance_Assists_relative', 'Chance_Contributions_relative', 'Chances_relative',
                          'Cycle_Fcheck_Contributions_relative', 'Rush_Contributions_relative', 'Zone_Entries_relative']),
    html.Div(id = 'stats-error', style = {"color" : "blue"}),
    dbc.Button('Submit', id = 'my_button', n_clicks = 0),
    dcc.Graph(id = "graph1")
])

### Update Player Dropdown Menus Based On Position

@callback(
    Output('players', 'options'),
    Output('players', 'value'),
    Input('positions', 'value'),
    Input('seasons', 'value')
)

def update_player(select_pos, select_season):
    filtered_possea = player_melt[(player_melt['Position'] == select_pos) & (player_melt['Season'] == select_season)]
    filtered_players = sorted(filtered_possea['Player'].unique())
    options = [{'label': player, 'value': player} for player in filtered_players]
    value = [filtered_players[0]] if len(filtered_players) > 0 else []
    return options, value
### Add Controls for interaction

@callback(
    Output(component_id = "graph1", component_property = "figure"),
    Output(component_id = "stats-error", component_property = "children"),
    Input("my_button", "n_clicks"),
    State(component_id = "players", component_property = "value"),
    State(component_id = "seasons", component_property = "value"),
    State(component_id = "positions", component_property = "value"),
    State(component_id = "stats", component_property = "value"),
    prevent_initial_call = True
)

def update_graph(button, players, season, position, stats):
    if stats is None or len(stats) < 3:
        return {}, 'Please Choose At Least 3 Stats'

    filtered_df = player_melt[
        (player_melt['Player'].isin(players)) &
        (player_melt['Season'] == season) &
        (player_melt['Position'] == position) &
        (player_melt['category'].isin(stats))]
    fig = px.line_polar(filtered_df, r = 'relative', theta = 'category', color = 'Player', line_close = True,
                        hover_name = 'Player', markers = True, direction = 'clockwise', start_angle = 180, render_mode = 'svg')
    return fig, ""

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)
#%%
