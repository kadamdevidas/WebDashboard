#### This version 5 contains all the details for five sheets PM01, PM02, Breakdown Maintenance,
#### Jobs Hold for Shutdown, Vibration Monitoring.

from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import pandas as pd
import plotly.graph_objects as go
from flask import Flask

# Create the Flask server
server = Flask(__name__)

# Create the Dash app and pass the Flask server
app = Dash(__name__, server=server, suppress_callback_exceptions=True)

# Load the Merged MPR Excel file
file_path = "./Merged_MPR.xlsx"

# Load data from each sheet
pm02_df = pd.read_excel(file_path, sheet_name='PM02')
pm01_df = pd.read_excel(file_path, sheet_name='PM01')
breakdown_df = pd.read_excel(file_path, sheet_name='Breakdown Maintenance')
shutdown_df = pd.read_excel(file_path, sheet_name='Jobs Hold for Shutdown')
vibration_df = pd.read_excel(file_path, sheet_name='Vibration Monitoring')

# Define styles for the sheet buttons (aligned properly)
button_style = {
    'display': 'inline-block',
    'width': '200px',
    'height': '100px',
    'lineHeight': '100px',
    'textAlign': 'center',
    'margin': '10px',
    'color': 'white',
    'fontSize': '20px',
    'borderRadius': '10px',
    'cursor': 'pointer',
    'fontFamily': 'Arial',
    'boxShadow': '2px 2px 5px rgba(0,0,0,0.2)',
    'verticalAlign': 'top'
}

# Define initial styles for the month buttons
selected_month_button_style = {
    'margin': '5px',
    'color': 'white',
    'backgroundColor': 'green',  # Green for selected buttons
    'border': 'none',
    'padding': '10px 20px',
    'cursor': 'pointer',
    'fontWeight': 'bold',
    'borderRadius': '5px',
}

deselected_month_button_style = {
    'margin': '5px',
    'color': 'white',
    'backgroundColor': 'black',  # Black for deselected buttons
    'border': 'none',
    'padding': '10px 20px',
    'cursor': 'pointer',
    'fontWeight': 'bold',
    'borderRadius': '5px',
}

# Define app layout
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content')
])

# Home page layout
home_layout = html.Div(style={'textAlign': 'center', 'padding': '50px'}, children=[
    html.H1('Mechanical Department Dashboard', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#161D6F', 'color': 'white'}),
    html.Div("Select a sheet to view details", style={'fontSize': '24px', 'padding': '20px'}),

    # First row for PM01, PM02, and Breakdown buttons
    html.Div(style={'textAlign': 'center', 'display': 'flex', 'justifyContent': 'center'}, children=[
        dcc.Link(html.Div("PM01", style={**button_style, 'backgroundColor': '#243642'}), href='/pm01'),
        dcc.Link(html.Div("PM02", style={**button_style, 'backgroundColor': '#387478'}), href='/pm02'),
        dcc.Link(html.Div("Breakdown", style={**button_style, 'backgroundColor': '#629584'}), href='/breakdown')
    ]),

    # Second row for Shutdown Jobs and Vibration Monitoring buttons
    html.Div(style={'textAlign': 'center', 'display': 'flex', 'justifyContent': 'center'}, children=[
        dcc.Link(html.Div("Shutdown Jobs", style={**button_style, 'backgroundColor': '#A45E60'}), href='/shutdown'),
        dcc.Link(html.Div("Vibration Monitoring", style={**button_style, 'backgroundColor': '#4682B4'}), href='/vibration')
    ]),
])

# Layout for PM01, PM02, Breakdown, Shutdown, and Vibration Monitoring
def create_sheet_layout(sheet_name, id_prefix):
    return html.Div(style={'backgroundColor': '#f9f9f9', 'font-family': 'Arial'}, children=[
        dcc.Link('Back to Home', href='/', style={'display': 'block', 'textAlign': 'center', 'padding': '20px', 'fontSize': '20px'}),
        html.H2(f'{sheet_name}', style={'textAlign': 'center', 'padding': '20px', 'backgroundColor': '#50e3c2', 'color': 'white'}),

        # Month filter buttons
        html.Div(id=f'{id_prefix}-month-buttons', style={'textAlign': 'center', 'padding': '10px'}),
        html.Div(id=f'{id_prefix}-graphs', style={'display': 'flex', 'flex-wrap': 'wrap', 'justify-content': 'center'}),
        html.Div(id=f'{id_prefix}-tables', style={'padding': '20px', 'textAlign': 'center'})
    ])

# Create layouts for each sheet
pm01_layout = create_sheet_layout('PM01 Unplanned Maintenance', 'pm01')
pm02_layout = create_sheet_layout('PM02 Planned Maintenance', 'pm02')
breakdown_layout = create_sheet_layout('Breakdown Maintenance', 'breakdown')
shutdown_layout = create_sheet_layout('Jobs Hold for Shutdown', 'shutdown')
vibration_layout = create_sheet_layout('Vibration Monitoring', 'vibration')

# Callback to dynamically update the page content
@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/pm01':
        return pm01_layout
    elif pathname == '/pm02':
        return pm02_layout
    elif pathname == '/breakdown':
        return breakdown_layout
    elif pathname == '/shutdown':
        return shutdown_layout
    elif pathname == '/vibration':
        return vibration_layout
    else:
        return home_layout

# Callback to dynamically generate month buttons for each sheet
def generate_month_buttons(sheet_df, prefix):
    months = sheet_df['Month'].unique()
    buttons = []
    for month in months:
        buttons.append(html.Button(month, id=f'btn-{prefix}-{month}', n_clicks=1, style=selected_month_button_style))
    return buttons

# Generate PM01 month buttons
@app.callback(
    Output('pm01-month-buttons', 'children'),
    Input('url', 'pathname')
)
def generate_pm01_month_buttons(pathname):
    if pathname == '/pm01':
        return generate_month_buttons(pm01_df, 'pm01')

# Generate PM02 month buttons
@app.callback(
    Output('pm02-month-buttons', 'children'),
    Input('url', 'pathname')
)
def generate_pm02_month_buttons(pathname):
    if pathname == '/pm02':
        return generate_month_buttons(pm02_df, 'pm02')

# Generate Breakdown month buttons
@app.callback(
    Output('breakdown-month-buttons', 'children'),
    Input('url', 'pathname')
)
def generate_breakdown_month_buttons(pathname):
    if pathname == '/breakdown':
        return generate_month_buttons(breakdown_df, 'breakdown')

# Generate Shutdown month buttons
@app.callback(
    Output('shutdown-month-buttons', 'children'),
    Input('url', 'pathname')
)
def generate_shutdown_month_buttons(pathname):
    if pathname == '/shutdown':
        return generate_month_buttons(shutdown_df, 'shutdown')

# Generate Vibration Monitoring month buttons
@app.callback(
    Output('vibration-month-buttons', 'children'),
    Input('url', 'pathname')
)
def generate_vibration_month_buttons(pathname):
    if pathname == '/vibration':
        return generate_month_buttons(vibration_df, 'vibration')

# Update graphs for PM01 and PM02
def update_pm_graphs(sheet_df, prefix):
    @app.callback(
        [Output(f'btn-{prefix}-{month}', 'style') for month in sheet_df['Month'].unique()] + 
        [Output(f'{prefix}-graphs', 'children')],
        [Input(f'btn-{prefix}-{month}', 'n_clicks') for month in sheet_df['Month'].unique()]
    )
    def update_sheet_graphs(*clicked_buttons):
        months_clicked = []
        styles = []
        graphs = []

        for i, month in enumerate(sheet_df['Month'].unique()):
            if clicked_buttons[i] % 2 == 1:  # Selected
                months_clicked.append(month)
                styles.append(selected_month_button_style)  # Keep green background
            else:  # Deselected
                styles.append(deselected_month_button_style)  # Turn to black background

        # Generate graphs for the selected months
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
        for i, month in enumerate(months_clicked):
            month_data = sheet_df[sheet_df['Month'] == month]
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=month_data['Plant'], y=month_data['Planned'], name='Planned',
                marker_color=colors[i % len(colors)], text=month_data['Planned'], textposition='auto'
            ))
            fig.add_trace(go.Bar(
                x=month_data['Plant'],                y=month_data['Executed'], name='Executed',
                marker_color=colors[(i + 1) % len(colors)], text=month_data['Executed'], textposition='auto'
            ))
            fig.update_layout(
                title=f'{prefix.upper()}: Planned vs Executed Jobs ({month})',
                barmode='group',
                height=350,
                width=350,
                plot_bgcolor='#f9f9f9'
            )
            graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

        return styles + [graphs]

# Update Breakdown graphs and tables
@app.callback(
    [Output(f'btn-breakdown-{month}', 'style') for month in breakdown_df['Month'].unique()] +
    [Output('breakdown-graphs', 'children'), Output('breakdown-tables', 'children')],
    [Input(f'btn-breakdown-{month}', 'n_clicks') for month in breakdown_df['Month'].unique()]
)
def update_breakdown_maintenance(*clicked_buttons):
    months_clicked = []
    styles = []
    graphs = []
    tables = []

    for i, month in enumerate(breakdown_df['Month'].unique()):
        if clicked_buttons[i] % 2 == 1:  # Selected
            months_clicked.append(month)
            styles.append(selected_month_button_style)  # Keep green background
        else:  # Deselected
            styles.append(deselected_month_button_style)  # Turn to black background

    # Generate graphs and tables for the selected months
    for i, month in enumerate(months_clicked):
        month_data = breakdown_df[breakdown_df['Month'] == month]

        # Generate Bar Chart for Breakdown Jobs
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Breakdown Jobs'], name='Breakdown Jobs',
            marker_color='#ff7f0e', text=month_data['No. of Breakdown Jobs'], textposition='auto'
        ))
        fig.update_layout(
            title=f'Breakdown Jobs ({month})',
            barmode='group',
            height=350,
            width=350,
            plot_bgcolor='#f9f9f9'
        )
        graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

        # Generate Table for Short Description with better styling
        table_rows = []
        for j, row in month_data.iterrows():
            table_rows.append(html.Tr([
                html.Td(j + 1, style={'padding': '10px', 'border': '1px solid black'}),
                html.Td(row['Plant'], style={'padding': '10px', 'border': '1px solid black'}),
                html.Td(row['Short description of the job'] if not pd.isna(row['Short description of the job']) else '', 
                        style={'padding': '10px', 'border': '1px solid black'})
            ], style={'backgroundColor': '#f9f9f9' if j % 2 == 0 else '#e0e0e0'}))  # Alternating row colors

        tables.append(html.Div([
            html.H3(f'{month} Breakdown Maintenance Jobs'),
            html.Table([
                html.Thead(html.Tr([
                    html.Th('Serial Number', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'}),
                    html.Th('Plant', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'}),
                    html.Th('Short description of the job', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'})
                ])),
                html.Tbody(table_rows)
            ], style={'width': '80%', 'margin': '20px auto', 'border': '1px solid black', 'borderCollapse': 'collapse',
                      'textAlign': 'center', 'borderSpacing': '0px', 'fontFamily': 'Arial'})
        ], style={'padding': '20px'}))

    return styles + [graphs, tables]

# Update Shutdown Jobs graphs and tables
@app.callback(
    [Output(f'btn-shutdown-{month}', 'style') for month in shutdown_df['Month'].unique()] +
    [Output('shutdown-graphs', 'children'), Output('shutdown-tables', 'children')],
    [Input(f'btn-shutdown-{month}', 'n_clicks') for month in shutdown_df['Month'].unique()]
)
def update_shutdown_jobs(*clicked_buttons):
    months_clicked = []
    styles = []
    graphs = []
    tables = []

    for i, month in enumerate(shutdown_df['Month'].unique()):
        if clicked_buttons[i] % 2 == 1:  # Selected
            months_clicked.append(month)
            styles.append(selected_month_button_style)  # Keep green background
        else:  # Deselected
            styles.append(deselected_month_button_style)  # Turn to black background

    # Generate graphs and tables for the selected months
    for i, month in enumerate(months_clicked):
        month_data = shutdown_df[shutdown_df['Month'] == month]

        # Generate Bar Chart for Jobs Hold for Shutdown
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Jobs hold'], name='Jobs Hold',
            marker_color='#ff7f0e', text=month_data['No. of Jobs hold'], textposition='auto'
        ))
        fig.update_layout(
            title=f'Shutdown Jobs ({month})',
            barmode='group',
            height=350,
            width=350,
            plot_bgcolor='#f9f9f9'
        )
        graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

        # Generate Table for Short Description with better styling
        table_rows = []
        for j, row in month_data.iterrows():
            table_rows.append(html.Tr([
                html.Td(j + 1, style={'padding': '10px', 'border': '1px solid black'}),
                html.Td(row['Plant'], style={'padding': '10px', 'border': '1px solid black'}),
                html.Td(row['Short description of the job'] if not pd.isna(row['Short description of the job']) else '', 
                        style={'padding': '10px', 'border': '1px solid black'})
            ], style={'backgroundColor': '#f9f9f9' if j % 2 == 0 else '#e0e0e0'}))  # Alternating row colors

        tables.append(html.Div([
            html.H3(f'{month} Jobs Hold for Shutdown'),
            html.Table([
                html.Thead(html.Tr([
                    html.Th('Serial Number', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'}),
                    html.Th('Plant', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'}),
                    html.Th('Short description of the job', style={'padding': '10px', 'border': '1px solid black', 'backgroundColor': '#d3a15d', 'color': 'white'})
                ])),
                html.Tbody(table_rows)
            ], style={'width': '80%', 'margin': '20px auto', 'border': '1px solid black', 'borderCollapse': 'collapse',
                      'textAlign': 'center', 'borderSpacing': '0px', 'fontFamily': 'Arial'})
        ], style={'padding': '20px'}))

    return styles + [graphs, tables]

# Update Vibration Monitoring graphs
@app.callback(
    [Output(f'btn-vibration-{month}', 'style') for month in vibration_df['Month'].unique()] +
    [Output('vibration-graphs', 'children')],
    [Input(f'btn-vibration-{month}', 'n_clicks') for month in vibration_df['Month'].unique()]
)
def update_vibration_monitoring(*clicked_buttons):
    months_clicked = []
    styles = []
    graphs = []

    for i, month in enumerate(vibration_df['Month'].unique()):
        if clicked_buttons[i] % 2 == 1:  # Selected
            months_clicked.append(month)
            styles.append(selected_month_button_style)  # Keep green background
        else:  # Deselected
            styles.append(deselected_month_button_style)  # Turn to black background

    # Generate graphs for the selected months
    for i, month in enumerate(months_clicked):
        month_data = vibration_df[vibration_df['Month'] == month]

        # Generate Bar Chart for Vibration Monitoring Jobs
        fig = go.Figure()

        # Left Y-axis (Equipment-related counts)
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Equipment Scheduled'], name='Scheduled Equipment',
            marker_color='#1f77b4', text=month_data['No. of Equipment Scheduled'], textposition='auto'
        ))
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Equipment Monitored (Executed)'], name='Monitored Equipment',
            marker_color='#ff7f0e', text=month_data['No. of Equipment Monitored (Executed)'], textposition='auto'
        ))
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Equipment With Normal Health'], name='Normal Health Equipment',
            marker_color='#2ca02c', text=month_data['No. of Equipment With Normal Health'], textposition='auto'
        ))
        fig.add_trace(go.Bar(
            x=month_data['Plant'], y=month_data['No. of Equipment With Critical Health'], name='Critical Health Equipment',
            marker_color='#d62728', text=month_data['No. of Equipment With Critical Health'], textposition='auto'
        ))

        # Right Y-axis (Percentage-related data)
        fig.add_trace(go.Scatter(
            x=month_data['Plant'], y=month_data['Health Index'] * 100, name='Health Index (%)',
            marker_color='#9467bd', yaxis='y2', mode='lines+markers', text=month_data['Health Index'] * 100, textposition='top center'
        ))
        fig.add_trace(go.Scatter(
            x=month_data['Plant'], y=month_data['Critical Index'] * 100, name='Critical Index (%)',
            marker_color='#e8bd13', yaxis='y2', mode='lines+markers', text=month_data['Critical Index'] * 100, textposition='top center'
        ))

        # Update layout for dual-axis graph
        fig.update_layout(
            title=f'Vibration Monitoring ({month})',
            barmode='group',
            height=400,
            width=700,
            plot_bgcolor='#f9f9f9',
            xaxis=dict(title='Plant'),
            yaxis=dict(title='Equipment Count'),
            yaxis2=dict(title='Percentage (%)', overlaying='y', side='right'),
            legend=dict(x=1.05, y=1),
            margin=dict(l=40, r=40, t=40, b=40)
        )

        graphs.append(html.Div(dcc.Graph(figure=fig), style={'display': 'inline-block', 'margin': '10px'}))

    return styles + [graphs]

# Generate the month buttons for PM01, PM02, Breakdown, Shutdown, and Vibration Monitoring
update_pm_graphs(pm01_df, 'pm01')
update_pm_graphs(pm02_df, 'pm02')

# Main entry point
if __name__ == "__main__":
    app.run_server(debug=True)


