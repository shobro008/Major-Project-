import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import calendar

# ==========================================
# 1. LOAD AND PREP THE DATA
# ==========================================
print("Loading data... (This might take a minute)")
# Replace 'your_data.xlsx' with your actual file name
# Replace 'your_data.csv' with your actual file name
df = pd.read_excel('forecast.xlsx')
print(df.columns)

# Change 'Datetime' to the actual name of your time column in Excel
time_col = 'time'

df[time_col] = pd.to_datetime(df[time_col], format='%Y%m%d:%H%M')
# YEH NAYI LINE ADD KAREIN: UTC ko IST mein shift karne ke liye (+ 5.5 hours)
df[time_col] = df[time_col] + pd.Timedelta(hours=5, minutes=30)
df.set_index(time_col, inplace=True)

# ==========================================
# 2. DATA AGGREGATION (LEARNING THE 5-YEAR AVERAGES)
# ==========================================
print("Calculating 5-year averages for the forecast model...")

# A. Monthly Averages (Jan - Dec)
# Group by the 'month' of the index (1 to 12) and take the mean of all 5 years
monthly_avg = df.groupby(df.index.month).mean().reset_index()
# Map month numbers (1-12) to actual names (Jan, Feb, etc.)
monthly_avg[time_col] = monthly_avg[time_col].apply(lambda x: calendar.month_abbr[x])

# B. Hourly Averages (00:00 - 23:00)
# Group by the 'hour' of the index to see an average day's performance
hourly_avg = df.groupby(df.index.hour).mean().reset_index()

# ==========================================
# 3. BUILD THE DASHBOARD (INTERACTIVE PLOTS)
# ==========================================
print("Generating interactive dashboard...")

# Create a figure with subplots (2 rows, 2 columns)
fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=(
        "1. Expected PV Power per Month (5-Year Average)",
        "2. Average Daily Power Profile (Hour-by-Hour)",
        "3. Average Monthly Solar Irradiance",
        "4. Average Monthly Temperature"
    ),
    vertical_spacing=0.15,
    horizontal_spacing=0.1
)

# --- PLOT 1: Average Monthly PV Power (Top Left) ---
# This shows "Jan mei itna forecast"
fig.add_trace(
    go.Bar(x=monthly_avg[time_col], y=monthly_avg['P'],
           name='Avg PV Power (W)', marker_color='#2ca02c'),
    row=1, col=1
)

# --- PLOT 2: Average Hourly Daily Profile (Top Right) ---
fig.add_trace(
    go.Scatter(x=hourly_avg[time_col], y=hourly_avg['P'],
               mode='lines+markers', name='Avg Hourly Power (W)',
               line=dict(color='#d62728', width=3), fill='tozeroy'),
    row=1, col=2
)

# --- PLOT 3: Average Monthly Irradiance (Bottom Left) ---
fig.add_trace(
    go.Bar(x=monthly_avg[time_col], y=monthly_avg['Gb(i)'],
           name='Direct Irradiance', marker_color='#ff7f0e'),
    row=2, col=1
)
fig.add_trace(
    go.Bar(x=monthly_avg[time_col], y=monthly_avg['Gd(i)'],
           name='Diffuse Irradiance', marker_color='#1f77b4'),
    row=2, col=1
)
# Stack the irradiance bars on top of each other
fig.update_layout(barmode='stack')

# --- PLOT 4: Average Monthly Temperature (Bottom Right) ---
fig.add_trace(
    go.Scatter(x=monthly_avg[time_col], y=monthly_avg['T2m'],
               mode='lines+markers', name='Avg Temp (°C)',
               line=dict(color='#9467bd', width=3)),
    row=2, col=2
)

# ==========================================
# 4. DASHBOARD STYLING AND EXPORT
# ==========================================
# Update axis labels and layout
fig.update_layout(
    title_text="PV Forecasting: 5-Year Historical Average Baseline Model",
    title_font_size=24,
    height=900,
    width=1400,
    showlegend=True,
    template="plotly_white",
    hovermode="x unified" # Shows all data for a month when you hover over it
)

# Set specific axis titles and formatting
fig.update_yaxes(title_text="Power (W)", row=1, col=1)
fig.update_xaxes(title_text="Hour of Day (0-23)", tickmode='linear', tick0=0, dtick=2, row=1, col=2)
fig.update_yaxes(title_text="Power (W)", row=1, col=2)
fig.update_yaxes(title_text="Irradiance (W/m²)", row=2, col=1)
fig.update_yaxes(title_text="Temperature (°C)", row=2, col=2)

# Save the dashboard as an interactive HTML file
output_file = "PV_Average_Forecast_Dashboard.html"
fig.write_html(output_file)
print(f"Dashboard successfully generated! Open '{output_file}' in your web browser.")