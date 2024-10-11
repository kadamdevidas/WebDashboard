#r"/Users/omkarkadam/Documents/Automate with Python/Merged_MPR.xlsx"

import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.graph_objects as go

# Load the Merged MPR Excel file
file_path = "./Merged_MPR.xlsx"


# Load data from each sheet
pm02_df = pd.read_excel(file_path, sheet_name='PM02')
pm01_df = pd.read_excel(file_path, sheet_name='PM01')
breakdown_df = pd.read_excel(file_path, sheet_name='Breakdown Maintenance')

# Initialize the Dash app
app = dash.Dash(__name__)

# Define app layout
app.layout = html.Div(style={'backgroundColor': '#f9f9f9', 'font-family': 'Arial'}, children=[
    html.H1(children='MPR Dashboard', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#4a90e2', 'color': 'white'}),

    # Dropdown for selecting months (multi-select)
    html.Div([
        html.Label('Select Months:', style={'font-weight': 'bold', 'padding': '10px'}),
        dcc.Dropdown(
            id='month-filter',
            options=[{'label': month, 'value': month} for month in pm02_df['Month'].unique()],
            value=pm02_df['Month'].unique(),
            multi=True,
            style={'width': '60%', 'margin': 'auto'}
        )
    ], style={'padding': '20px', 'textAlign': 'center', 'backgroundColor': '#eef3f7'}),

    # Section for PM01 charts
    html.H2('PM01: Planned vs Executed Jobs', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#50e3c2', 'color': 'white'}),
    html.Div(id='pm01-graphs', style={'display': 'flex', 'flex-wrap': 'wrap', 'justify-content': 'center'}),

    # Section for PM02 charts
    html.H2('PM02: Planned vs Executed Jobs', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#7ed321', 'color': 'white'}),
    html.Div(id='pm02-graphs', style={'display': 'flex', 'flex-wrap': 'wrap', 'justify-content': 'center'}),

    # Breakdown Maintenance Chart
    html.H2('Breakdown Jobs Distribution', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#f5a623', 'color': 'white'}),
    dcc.Graph(id='breakdown-graph', style={'padding': '20px'})
])

# Callback to update graphs based on selected months
@app.callback(
    [Output('pm01-graphs', 'children'),
     Output('pm02-graphs', 'children'),
     Output('breakdown-graph', 'figure')],
    [Input('month-filter', 'value')]
)
def update_graphs(selected_months):
    # Filter data based on selected months
    filtered_pm01 = pm01_df[pm01_df['Month'].isin(selected_months)]
    filtered_pm02 = pm02_df[pm02_df['Month'].isin(selected_months)]
    filtered_breakdown = breakdown_df[breakdown_df['Month'].isin(selected_months)]
    
    # Colors for months
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

    # PM01 Graphs (compact, row-wise)
    pm01_graphs = []
    for i, month in enumerate(selected_months):
        month_data = filtered_pm01[filtered_pm01['Month'] == month]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['Planned'], name='Planned', 
            marker_color=colors[i % len(colors)], text=month_data['Planned'], textposition='auto'
        ))
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['Executed'], name='Executed', 
            marker_color=colors[(i + 1) % len(colors)], text=month_data['Executed'], textposition='auto'
        ))
        fig.update_layout(
            title=f'PM01: Planned vs Executed Jobs ({month})', barmode='group', height=350, width=350,
            plot_bgcolor='#f9f9f9'
        )
        pm01_graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

    # PM02 Graphs (compact, row-wise, start new row after PM01)
    pm02_graphs = []
    for i, month in enumerate(selected_months):
        month_data = filtered_pm02[filtered_pm02['Month'] == month]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['Planned'], name='Planned', 
            marker_color=colors[i % len(colors)], text=month_data['Planned'], textposition='auto'
        ))
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['Executed'], name='Executed', 
            marker_color=colors[(i + 1) % len(colors)], text=month_data['Executed'], textposition='auto'
        ))
        fig.update_layout(
            title=f'PM02: Planned vs Executed Jobs ({month})', barmode='group', height=350, width=350,
            plot_bgcolor='#f9f9f9'
        )
        pm02_graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

    # Breakdown Maintenance Figure: Breakdown Jobs per Plant
    breakdown_fig = go.Figure(data=[go.Pie(
        labels=filtered_breakdown['Plant'], values=filtered_breakdown['No. of Breakdown Jobs'], hole=.3
    )])
    breakdown_fig.update_layout(
        title='Breakdown Jobs Distribution', height=400, plot_bgcolor='#f9f9f9'
    )

    return pm01_graphs, pm02_graphs, breakdown_fig

# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)
