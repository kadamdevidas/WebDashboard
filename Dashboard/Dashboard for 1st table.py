



import pandas as pd
from dash import Dash, dcc, html, Input, Output
import plotly.express as px

# Load the merged Excel file (considering only the first table)
file_path =r'/Users/omkarkadam/Documents/Automate with Python//Merged_MPR.xlsx'

# Table 1 has range [0:7, 0:4] (rows 0 to 6, columns 0 to 3, row 0 is header)
merged_table_1 = pd.read_excel(file_path, sheet_name='Merged Data', header=0, skiprows=0, nrows=7, usecols="A:D")

# Create the Dash app
app = Dash(__name__)

# Get unique months and plants for dropdowns
unique_months = merged_table_1['Month'].dropna().unique()
unique_plants = merged_table_1['Plant'].dropna().unique()

# Layout of the app
app.layout = html.Div([
    html.H1("MPR Dashboard (Table 1)", style={'text-align': 'center'}),

    # Dropdown to filter by month
    html.Div([
        html.Label('Select Month:'),
        dcc.Dropdown(
            id='month-dropdown',
            options=[{'label': month, 'value': month} for month in unique_months],
            value=unique_months[0],  # Default value
            clearable=False
        ),
    ], style={'width': '40%', 'display': 'inline-block'}),

    # Dropdown to filter by plant
    html.Div([
        html.Label('Select Plant:'),
        dcc.Dropdown(
            id='plant-dropdown',
            options=[{'label': plant, 'value': plant} for plant in unique_plants],
            value=unique_plants[0],  # Default value
            clearable=False
        ),
    ], style={'width': '40%', 'display': 'inline-block', 'margin-left': '20px'}),

    # Graph for Planned vs Executed
    html.Div([
        html.H3('Planned vs Executed'),
        dcc.Graph(id='table-1-graph'),
    ]),
])

# Callback to update the graph based on selected month and plant
@app.callback(
    Output('table-1-graph', 'figure'),
    [Input('month-dropdown', 'value'),
     Input('plant-dropdown', 'value')]
)
def update_graph(selected_month, selected_plant):
    # Filter the data for the selected month and plant
    filtered_table_1 = merged_table_1[(merged_table_1['Month'] == selected_month) & (merged_table_1['Plant'] == selected_plant)]

    # Create a bar chart showing Planned vs Executed
    fig = px.bar(filtered_table_1, x='Plant', y=['Planned', 'Executed'], barmode='group',
                 title=f"Planned vs Executed for {selected_plant} in {selected_month}")

    return fig

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)

