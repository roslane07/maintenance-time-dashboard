"""
COMPLETE Maintenance Time Analysis Dashboard
===========================================

A comprehensive Streamlit-based web interface that combines:
- Excel file upload and processing
- Interactive charts and visualizations for each day/sheet
- Real-time analysis and insights
- Advanced time calculations and duration analysis
- Donut charts and improvement recommendations
- Comprehensive reporting and data export
- All original analyser functionality integrated

Features:
✅ Excel file upload and processing
✅ Multi-sheet/day analysis
✅ Interactive visualizations (bar, donut, timeline charts)
✅ Time duration calculations
✅ Category mapping and color coding
✅ Efficiency analysis and recommendations
✅ Time waste identification
✅ Incidental activity analysis
✅ Comprehensive reporting
✅ Data export capabilities
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64
from datetime import datetime, timedelta, time as dtime
from io import BytesIO
import numpy as np
import re
import os
from fpdf import FPDF
import plotly.io as pio
from PIL import Image
import tempfile

# Page configuration
st.set_page_config(
    page_title="Complete Maintenance Time Analysis Dashboard",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .success-metric {
        border-left-color: #28a745;
    }
    .warning-metric {
        border-left-color: #ffc107;
    }
    .danger-metric {
        border-left-color: #dc3545;
    }
    .day-header {
        background-color: #1f2937; /* dark, simple, classy */
        color: #e5e7eb; /* ensure readable text */
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #2a2a2a; /* subtle outline */
    }
    .insight-card {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border-left: 4px solid #ffc107;
    }
    .improvement-card {
        background-color: #1f2937; /* dark slate */
        color: #e5e7eb; /* light text */
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border-left: 4px solid #1f77b4; /* brand blue */
    }
    /* Light gray inline highlight for important tokens */
    .hl {
        background-color: #e5e7eb; /* light gray */
        color: #111827;            /* near-black for contrast */
        padding: 0 4px;
        border-radius: 4px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        white-space: pre-wrap;
        background-color: #1e1e1e; /* darker base for dark theme */
        color: #d0d0d0;
        border-radius: 6px 6px 0 0;
        border: 1px solid #2a2a2a;
        gap: 6px;
        padding: 8px 12px;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #2a2f3a;
        color: #ffffff;
        border-color: #3a3f4a;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1f77b4;
        color: #ffffff;
        border-color: #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# CORE CONFIGURATION AND MAPPINGS
# ============================================================================

# Activity-code → category map
CATEGORY_MAP = {
    "TOT": "Wrench Time",
    "ISU": "Incidental", "ICL": "Incidental", "ITR": "Incidental",
    "IMG": "Incidental", "IBR": "Incidental", "ISL": "Incidental",
    "WID": "Waiting", "WWG": "Waiting", "WSR": "Waiting",
    "WOT": "Waiting", "WRW": "Waiting", "ISC": "Waiting"
}

COLORS = {
    "Wrench Time": "#2E8B57",  # Sea Green
    "Incidental": "#FF8C00",   # Dark Orange
    "Waiting": "#DC143C",      # Crimson Red
    "Other": "#6C757D"         # Gray
}

# ============================================================================
# CORE FUNCTIONS (from analyser.py)
# ============================================================================

def calculate_time_duration(times):
    """Calculate duration in minutes between time entries"""
    if len(times) < 2:
        return 3  # Default 3 minutes if only one entry
    
    # Convert time strings to datetime objects
    time_objects = []
    for time_str in times:
        if pd.notna(time_str) and str(time_str) != 'nan':
            try:
                if isinstance(time_str, str):
                    time_obj = pd.to_datetime(time_str).time()
                else:
                    time_obj = time_str
                time_objects.append(time_obj)
            except:
                continue
    
    if len(time_objects) < 2:
        return 3
    
    # Calculate duration in minutes
    start_time = time_objects[0]
    end_time = time_objects[-1]
    
    # Convert to datetime for calculation
    base_date = datetime(2024, 1, 1)
    start_dt = datetime.combine(base_date, start_time)
    end_dt = datetime.combine(base_date, end_time)
    
    duration = (end_dt - start_dt).total_seconds() / 60
    
    # If duration is 0 or negative, assume 3 minutes
    return max(duration, 3)

def extract_insights(data, category):
    """Extract insights from descriptions for a specific category"""
    category_data = data[data["Category"] == category]
    
    if category_data.empty:
        return []
    
    descriptions = category_data["Description"].str.lower().fillna("")
    
    insights = []
    
    if category == "Waiting":
        # Find waiting causes
        waiting_patterns = [
            r"wait(?:ing)?\s+(?:for|on)\s+([a-z\s]+)",
            r"wait(?:ing)?\s+([a-z\s]+)",
            r"([a-z\s]+)\s+delay",
            r"([a-z\s]+)\s+issue"
        ]
        
        for pattern in waiting_patterns:
            matches = descriptions.str.extract(pattern, expand=False)
            for match in matches.dropna():
                if len(match.strip()) > 2:
                    insights.append(match.strip())
    
    elif category == "Incidental":
        # Find incidental activities
        incidental_patterns = [
            r"(meeting|break|lunch|coffee|planning|discussion)",
            r"(admin|administrative|paperwork|documentation)",
            r"(training|learning|study)"
        ]
        
        for pattern in incidental_patterns:
            matches = descriptions.str.extract(pattern, expand=False)
            for match in matches.dropna():
                if len(match.strip()) > 2:
                    insights.append(match.strip())
    
    # Count occurrences and return top insights
    if insights:
        insight_counts = pd.Series(insights).value_counts()
        return insight_counts.head(5).to_dict()
    
    return []

def process_excel_data(uploaded_file):
    """Process uploaded Excel file and return processed data"""
    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        all_data = []
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Read the sheet
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                
                # Normalize and map column names to accept flexible headers
                # Required by user: "time ", " activity code", "  Detailed description of activity"
                def _norm(s):
                    return re.sub(r"\s+", " ", str(s)).strip().lower()

                col_map = {}
                for c in df.columns:
                    key = _norm(c)
                    if key == "time":
                        col_map[c] = "Time"
                    elif key == "activity code":
                        col_map[c] = "Activity Code"
                    elif key in ("detailed description of activity", "description"):
                        col_map[c] = "Description"

                df = df.rename(columns=col_map)

                required_cols = ['Time', 'Activity Code', 'Description']
                if not all(col in df.columns for col in required_cols):
                    st.warning(
                        f"Sheet '{sheet_name}' missing required columns (expected: time, activity code, detailed description of activity). Skipping..."
                    )
                    continue
                
                # Clean and process the data
                df = df[required_cols].copy()
                df = df.dropna(subset=['Time', 'Activity Code'])
                
                # Add sheet name for tracking
                df['Sheet'] = sheet_name
                
                # Categorize activities
                df['Category'] = df['Activity Code'].map(CATEGORY_MAP).fillna('Other')
                
                # Each line represents 3 minutes
                df['Duration'] = 3.0
                
                all_data.append(df)
                
            except Exception as e:
                st.warning(f"Error processing sheet '{sheet_name}': {str(e)}")
                continue
        
        if not all_data:
            st.error("No valid data found in any sheet!")
            return None
        
        # Combine all data
        combined_data = pd.concat(all_data, ignore_index=True)

        # Robustly normalize Time column to datetime.time objects
        def _to_time(val):
            if pd.isna(val):
                return pd.NA
            # Already a time object
            if isinstance(val, dtime):
                return val
            # pandas/py datetime
            if isinstance(val, (pd.Timestamp, datetime)):
                try:
                    return val.to_pydatetime().time() if isinstance(val, pd.Timestamp) else val.time()
                except Exception:
                    return pd.NA
            # Excel serial as number (fraction of a day)
            if isinstance(val, (int, float)):
                try:
                    t = (datetime(1899, 12, 30) + timedelta(days=float(val))).time()
                    return t
                except Exception:
                    return pd.NA
            # String parsing
            try:
                ts = pd.to_datetime(str(val), errors='coerce')
                if pd.isna(ts):
                    return pd.NA
                if isinstance(ts, pd.Timestamp):
                    return ts.to_pydatetime().time()
                return pd.NA
            except Exception:
                return pd.NA

        combined_data['Time'] = combined_data['Time'].apply(_to_time)
        combined_data = combined_data.dropna(subset=['Time'])
        
        return combined_data
        
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        return None

# ============================================================================
# DASHBOARD FUNCTIONS
# ============================================================================

def create_donut_chart(data, title, height=400):
    """Create a donut chart for category distribution"""
    if data.empty:
        return go.Figure()
    
    category_counts = data['Category'].value_counts()
    
    fig = go.Figure(data=[go.Pie(
        labels=category_counts.index,
        values=category_counts.values,
        hole=0.6,
        marker_colors=[COLORS.get(cat, '#6C757D') for cat in category_counts.index],
        textinfo='label+percent',
        textposition='outside'
    )])
    
    fig.update_layout(
        title=title,
        height=height,
        showlegend=True,
        margin=dict(t=50, b=50, l=50, r=50)
    )
    
    return fig

def create_hourly_analysis(data):
    """Create hourly time distribution analysis"""
    if data.empty:
        return go.Figure()
    
    # Convert time to hour
    data_copy = data.copy()
    data_copy['Hour'] = data_copy['Time'].apply(lambda t: t.hour if pd.notna(t) else pd.NA)
    data_copy = data_copy.dropna(subset=['Hour'])
    
    hourly_data = data_copy.groupby(['Hour', 'Category'])['Duration'].sum().reset_index()
    
    fig = px.bar(
        hourly_data,
        x='Hour',
        y='Duration',
        color='Category',
        color_discrete_map=COLORS,
        title="Hourly Time Distribution by Category",
        labels={'Hour': 'Hour of Day', 'Duration': 'Duration (minutes)'}
    )
    
    fig.update_layout(
        xaxis=dict(tickmode='linear', tick0=0, dtick=1),
        height=500
    )
    
    return fig

def create_activity_timeline(data):
    """Create activity timeline visualization"""
    if data.empty:
        return go.Figure()

    # Use all activities for the provided dataset (full day)
    timeline_data = data.copy()
    timeline_data['Duration'] = timeline_data['Duration'].fillna(3)

    # Build start and end datetimes using a constant date
    base_date = datetime(2024, 1, 1)
    timeline_data['Start'] = timeline_data['Time'].apply(lambda t: datetime.combine(base_date, t))
    timeline_data['End'] = timeline_data.apply(lambda r: r['Start'] + timedelta(minutes=float(r['Duration'])), axis=1)

    fig = px.timeline(
        timeline_data,
        x_start='Start',
        x_end='End',
        y='Description',
        color='Category',
        color_discrete_map=COLORS,
        title="Activity Timeline",
        labels={'Start': 'Start', 'End': 'End', 'Description': 'Activity', 'Category': 'Category'}
    )

    # Show time-only on axis and hover; do not display any date component
    fig.update_layout(height=500)
    fig.update_xaxes(tickformat="%H:%M", hoverformat="%H:%M")
    return fig

def create_daily_summary_charts(data):
    """Create comprehensive daily summary charts"""
    if data.empty:
        return None
    
    # Group by sheet (day) and category
    daily_summary = data.groupby(['Sheet', 'Category'])['Duration'].sum().reset_index()
    
    # Create subplots
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Daily Category Distribution', 'Daily Efficiency', 'Top Activities by Day', 'Category Trends'),
        specs=[[{"type": "pie"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "scatter"}]]
    )
    
    # 1. Daily Category Distribution (Pie)
    total_by_category = data.groupby('Category')['Duration'].sum()
    fig.add_trace(
        go.Pie(labels=total_by_category.index, values=total_by_category.values, name="Total"),
        row=1, col=1
    )
    
    # 2. Daily Efficiency (Bar)
    daily_efficiency = data.groupby('Sheet').apply(
        lambda x: (x[x['Category'] == 'Wrench Time']['Duration'].sum() / x['Duration'].sum()) * 100
    ).reset_index()
    daily_efficiency.columns = ['Day', 'Efficiency %']
    
    fig.add_trace(
        go.Bar(x=daily_efficiency['Day'], y=daily_efficiency['Efficiency %'], name="Efficiency %"),
        row=1, col=2
    )
    
    # 3. Top Activities by Day (Bar)
    top_activities = data.groupby(['Sheet', 'Category'])['Duration'].sum().reset_index()
    top_activities = top_activities.sort_values('Duration', ascending=False).head(10)
    
    fig.add_trace(
        go.Bar(x=top_activities['Sheet'], y=top_activities['Duration'], 
               color=top_activities['Category'], name="Top Activities"),
        row=2, col=1
    )
    
    # 4. Category Trends (Scatter)
    category_trends = data.groupby(['Sheet', 'Category'])['Duration'].sum().reset_index()
    
    for category in COLORS.keys():
        cat_data = category_trends[category_trends['Category'] == category]
        if not cat_data.empty:
            fig.add_trace(
                go.Scatter(x=cat_data['Sheet'], y=cat_data['Duration'], 
                          mode='lines+markers', name=category),
                row=2, col=2
            )
    
    fig.update_layout(height=800, showlegend=True)
    return fig

def create_incidental_analysis(data):
    """Create incidental activities analysis"""
    if data.empty:
        return go.Figure()
    
    incidental_data = data[data['Category'] == 'Incidental'].copy()
    
    if incidental_data.empty:
        return go.Figure()
    
    # Analyze incidental activities by description
    incidental_summary = incidental_data.groupby('Description')['Duration'].sum().sort_values(ascending=False).head(10)
    
    fig = px.bar(
        x=incidental_summary.values,
        y=incidental_summary.index,
        orientation='h',
        title="Top Incidental Activities by Duration",
        labels={'x': 'Duration (minutes)', 'y': 'Activity Description'},
        color=incidental_summary.values,
        color_continuous_scale='Oranges'
    )
    
    fig.update_layout(height=500)
    return fig

def create_improvement_opportunities_chart(data):
    """Create chart showing improvement opportunities"""
    if data.empty:
        return go.Figure()
    
    # Calculate potential improvements
    waiting_time = data[data['Category'] == 'Waiting']['Duration'].sum()
    incidental_time = data[data['Category'] == 'Incidental']['Duration'].sum()
    wrench_time = data[data['Category'] == 'Wrench Time']['Duration'].sum()
    
    # Create figure directly with go.Figure
    fig = go.Figure()
    
    # Add bars
    fig.add_trace(go.Bar(
        x=['Current Wrench Time'],
        y=[wrench_time],
        name='Current Wrench Time',
        marker_color='#10b981'
    ))
    
    fig.add_trace(go.Bar(
        x=['Potential Improvement'],
        y=[waiting_time + incidental_time * 0.5],
        name='Potential Improvement',
        marker_color='#60a5fa'
    ))
    
    fig.add_trace(go.Bar(
        x=['Remaining Time'],
        y=[incidental_time * 0.5],
        name='Remaining Time',
        marker_color='#6b7280'
    ))
    
    fig.update_layout(
        title="Wrench Time Improvement Opportunities",
        xaxis_title="Time Category",
        yaxis_title="Duration (minutes)",
        barmode='group',
        height=500
    )
    
    return fig

def generate_improvement_recommendations(data):
    """Generate improvement recommendations based on analysis"""
    recommendations = []
    
    # Calculate key metrics
    total_time = data['Duration'].sum()
    waiting_time = data[data['Category'] == 'Waiting']['Duration'].sum()
    incidental_time = data[data['Category'] == 'Incidental']['Duration'].sum()
    wrench_time = data[data['Category'] == 'Wrench Time']['Duration'].sum()
    
    # Calculate percentages
    if total_time > 0:
        waiting_pct = (waiting_time / total_time) * 100
        incidental_pct = (incidental_time / total_time) * 100
        wrench_pct = (wrench_time / total_time) * 100
        
        # Generate recommendations based on metrics
        if wrench_pct < 50:
            recommendations.append(f"Low wrench time efficiency ({wrench_pct:.1f}%). Target: Increase direct work time to at least 50%.")
        
        if waiting_pct > 15:
            recommendations.append(f"High waiting time ({waiting_pct:.1f}%). Action: Identify and address top delay causes.")
            
            # Get top waiting causes
            waiting_causes = data[data['Category'] == 'Waiting'].groupby('Description')['Duration'].sum()
            if not waiting_causes.empty:
                top_cause = waiting_causes.idxmax()
                recommendations.append(f"Focus on reducing '{top_cause}' which accounts for {waiting_causes.max()/waiting_time*100:.1f}% of waiting time.")
        
        if incidental_pct > 25:
            recommendations.append(f"High incidental time ({incidental_pct:.1f}%). Action: Review and optimize support activities.")
            
            # Get top incidental activities
            incidental_activities = data[data['Category'] == 'Incidental'].groupby('Description')['Duration'].sum()
            if not incidental_activities.empty:
                top_activity = incidental_activities.idxmax()
                recommendations.append(f"Review '{top_activity}' which takes up {incidental_activities.max()/incidental_time*100:.1f}% of incidental time.")
        
        # Daily variation analysis
        daily_efficiency = data.groupby('Sheet').apply(
            lambda x: (x[x['Category'] == 'Wrench Time']['Duration'].sum() / x['Duration'].sum()) * 100
        )
        if daily_efficiency.std() > 15:  # High variation
            recommendations.append("High daily variation in efficiency. Standardize work practices across all days.")
    
    if not recommendations:
        recommendations.append("No significant issues found. Continue monitoring and maintaining current performance.")
    
    return recommendations

def generate_excel_report(data, recommendations, figures=None):
    """Generate Excel report with all analysis"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = pd.DataFrame([
            {'Metric': 'Total Time (minutes)', 'Value': data['Duration'].sum()},
            {'Metric': 'Wrench Time %', 'Value': (data[data['Category'] == 'Wrench Time']['Duration'].sum() / data['Duration'].sum() * 100)},
            {'Metric': 'Waiting Time %', 'Value': (data[data['Category'] == 'Waiting']['Duration'].sum() / data['Duration'].sum() * 100)},
            {'Metric': 'Incidental Time %', 'Value': (data[data['Category'] == 'Incidental']['Duration'].sum() / data['Duration'].sum() * 100)}
        ])
        summary_data.to_excel(writer, sheet_name='Summary', index=False)
        
        # Raw data
        data.to_excel(writer, sheet_name='Raw Data', index=False)
        
        # Recommendations
        pd.DataFrame({'Recommendations': recommendations}).to_excel(
            writer, 
            sheet_name='Recommendations', 
            index=False
        )
        
        # Category analysis
        category_analysis = data.groupby('Category')['Duration'].agg(['sum', 'count']).reset_index()
        category_analysis.columns = ['Category', 'Total Minutes', 'Activity Count']
        category_analysis.to_excel(writer, sheet_name='Category Analysis', index=False)
        
        # Daily analysis
        daily_analysis = data.pivot_table(
            values='Duration',
            index='Sheet',
            columns='Category',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        daily_analysis.to_excel(writer, sheet_name='Daily Analysis', index=False)
    
    return output.getvalue()

def generate_pdf_report(data, charts, calculations, recommendations):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 12, "Maintenance Analysis Report", ln=True, align="C")
    pdf.ln(6)

    # Section: Summary & Calculations
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Summary & Calculations", ln=True)
    pdf.set_font("Arial", size=12)
    for key, value in calculations.items():
        pdf.cell(0, 8, f"{key}: {value}", ln=True)
    pdf.ln(4)

    # Section: Recommendations
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Recommendations", ln=True)
    pdf.set_font("Arial", size=12)
    for rec in recommendations:
        pdf.multi_cell(0, 8, f"- {rec}")
    pdf.ln(4)

    # Section: Charts
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Charts", ln=True)
    pdf.set_font("Arial", "I", 11)
    pdf.cell(0, 8, "Visualizations generated from your data:", ln=True)
    pdf.ln(2)

    # Add all charts as high-res images, smaller size for better layout
    for chart_name, fig in charts.items():
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
            pio.write_image(fig, tmpfile.name, format="png", width=700, height=350, scale=2)
            tmpfile.close()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 8, chart_name, ln=True)
            pdf.image(tmpfile.name, w=110)  # Smaller width for better structure
            pdf.ln(6)
        os.unlink(tmpfile.name)

    # --- Daily Analysis for Multi-sheet ---
    if data['Sheet'].nunique() > 1:
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 12, "Daily Analysis", ln=True, align="C")
        pdf.ln(6)
        for sheet in data['Sheet'].unique():
            sheet_data = data[data['Sheet'] == sheet]
            pdf.set_font("Arial", "B", 13)
            pdf.cell(0, 10, f"Day: {sheet}", ln=True)
            pdf.set_font("Arial", size=11)
            total_daily = sheet_data['Duration'].sum()
            wrench_daily = sheet_data[sheet_data['Category'] == 'Wrench Time']['Duration'].sum()
            efficiency_daily = (wrench_daily / total_daily) * 100 if total_daily > 0 else 0
            pdf.cell(0, 8, f"Total Time: {total_daily:.0f} min", ln=True)
            pdf.cell(0, 8, f"Wrench Time: {wrench_daily:.0f} min", ln=True)
            pdf.cell(0, 8, f"Daily Efficiency: {efficiency_daily:.1f}%", ln=True)
            pdf.ln(2)
            # Top activities
            top_activities = sheet_data.groupby('Description')['Duration'].sum().sort_values(ascending=False).head(5)
            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, "Top Activities:", ln=True)
            pdf.set_font("Arial", size=11)
            for activity, duration in top_activities.items():
                pdf.cell(0, 8, f"• {activity}: {duration:.0f} min", ln=True)
            pdf.ln(2)
            # Donut chart for the day
            donut_fig = create_donut_chart(sheet_data, f"{sheet} - Category Distribution", height=300)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                pio.write_image(donut_fig, tmpfile.name, format="png", width=700, height=300, scale=2)
                tmpfile.close()
                pdf.image(tmpfile.name, w=90)
                pdf.ln(2)
            os.unlink(tmpfile.name)
            # Timeline chart for the day
            timeline_fig = create_activity_timeline(sheet_data)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                pio.write_image(timeline_fig, tmpfile.name, format="png", width=700, height=300, scale=2)
                tmpfile.close()
                pdf.image(tmpfile.name, w=90)
                pdf.ln(4)
            os.unlink(tmpfile.name)
            pdf.ln(2)

    # Footer
    pdf.set_y(-25)
    pdf.set_font("Arial", "I", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 10, "Generated by Complete Maintenance Time Analysis Dashboard", align="C")

    return pdf.output(dest='S').encode('latin1')

# MAIN DASHBOARD INTERFACE
# ============================================================================

def main():
    st.markdown('<h1 class="main-header">Complete Maintenance Time Analysis Dashboard</h1>', unsafe_allow_html=True)
    st.markdown("---")

    # Sidebar for file upload
    st.sidebar.header("Upload Data")
    uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is None:
        st.info("Please upload an Excel file to begin analysis.")
        # Tool description for landing page (only before file upload, smaller style)
        st.markdown("""
        <div style="background-color:#23272f; color:#e5e7eb; padding:0.75rem 1.2rem; border-radius:0.4rem; margin-bottom:1.2rem; font-size:0.95rem;">
            <h4 style="margin-bottom:0.5rem;">What is this tool?</h4>
            <p style="margin-bottom:0.5rem;">
            <b>Complete Maintenance Time Analysis Dashboard</b> is an interactive analytics tool for maintenance teams and managers.<br>
            <ul style="margin-top:0.5rem; margin-bottom:0.5rem;">
                <li>Upload your maintenance Excel logs (single or multi-day).</li>
                <li>Instantly visualize time distribution, efficiency, waiting and incidental activities.</li>
                <li>Identify top time-wasters and improvement opportunities.</li>
                <li>Analyze daily and hourly trends, and get actionable recommendations.</li>
                <li>Export a comprehensive PDF report with all charts and insights.</li>
            </ul>
            <b>Purpose:</b> <i>Help you maximize wrench time, minimize delays, and optimize maintenance operations with data-driven insights.</i>
            </p>
        </div>
        """, unsafe_allow_html=True)
        # Pre-upload instructions
        st.markdown("""
        ### Expected Data Format:
        Your Excel file should contain these columns (extra spaces/casing are okay):
        - <span class="hl">time</span>
        - <span class="hl">activity code</span>
        - <span class="hl">detailed description of activity</span>

        The app will map these to internal columns: <span class="hl">Time, Activity Code, Description</span>.

        #### Activity Codes:
        - **Wrench Time**: <span class="hl">TOT</span>
        - **Incidental**: <span class="hl">ISU/ICL/ITR/IMG/IBR/ISL</span>
        - **Waiting**: <span class="hl">WID/WWG/WSR/WOT/WRW/ISC</span>
        """, unsafe_allow_html=True)
        return

    data = process_excel_data(uploaded_file)

    if data is None or data.empty:
        st.error("Failed to process data. Please check the file format and content.")
        return

    # --- Session State Initialization ---
    if 'report_ready' not in st.session_state:
        st.session_state.report_ready = False

    # --- Download Button Logic ---
    if st.sidebar.button('Prepare & Download Report'):
        with st.spinner('Generating PDF report...'):
            charts = {
                "Time Distribution Chart": create_donut_chart(data, "Overall Time Distribution"),
                "Efficiency Chart": create_improvement_opportunities_chart(data),
                "Hourly Analysis": create_hourly_analysis(data),
                "Time Waste Analysis": px.bar(
                    data[data['Category'] == 'Waiting'].groupby('Description')['Duration'].sum().sort_values(ascending=False).head(10),
                    orientation='h', title="Top Time Waste Causes"
                ),
                "Incidental Activities Chart": create_incidental_analysis(data),
                "Activity Timeline Chart": create_activity_timeline(data)
            }
            # Only add trends chart if multi-sheet
            if data['Sheet'].nunique() > 1:
                charts["Trends Chart"] = px.line(
                    data.groupby(['Sheet', 'Category'])['Duration'].sum().reset_index(),
                    x='Sheet', y='Duration', color='Category', title="Daily Category Trends"
                )
            calculations = {
                "Total Time (min)": data['Duration'].sum(),
                "Wrench Time %": f"{(data[data['Category']=='Wrench Time']['Duration'].sum()/data['Duration'].sum()*100):.1f}%",
            }
            recommendations = generate_improvement_recommendations(data)
            pdf_bytes = generate_pdf_report(data, charts, calculations, recommendations)
            st.sidebar.download_button(
                label="Download PDF Report",
                data=pdf_bytes,
                file_name="maintenance_analysis_report.pdf",
                mime="application/pdf"
            )

    # --- Main content tabs ---
    if data['Sheet'].nunique() > 1:
        tab_names = [
            "Overall Summary", "Efficiency", "Hourly Analysis", "Time Waste", 
            "Incidental Activities", "Trends", "Daily Analysis"
        ]
    else:
        tab_names = [
            "Overall Summary", "Efficiency", "Hourly Analysis", "Time Waste", 
            "Incidental Activities", "Daily Analysis"
        ]
    tabs = st.tabs(tab_names)
    # Assign tabs dynamically
    if data['Sheet'].nunique() > 1:
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = tabs
    else:
        tab1, tab2, tab3, tab4, tab5, tab7 = tabs

    # Tab 1: Overall Summary
    with tab1:
        st.header("Overall Performance Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(create_donut_chart(data, 'Time Distribution'), use_container_width=True)
        with col2:
            st.subheader("Key Metrics")
            total_time = data['Duration'].sum()
            wrench_time = data[data['Category'] == 'Wrench Time']['Duration'].sum()
            efficiency = (wrench_time / total_time * 100) if total_time > 0 else 0
            st.metric("Total Recorded Time", f"{total_time:,.0f} min")
            st.metric("Wrench Time Efficiency", f"{efficiency:.1f}%")
            st.metric("Total Wrench Time", f"{wrench_time:,.0f} min")

    # Tab 2: Efficiency
    with tab2:
        st.header("Wrench Time Efficiency Analysis")
        fig = create_improvement_opportunities_chart(data)
        st.plotly_chart(fig, use_container_width=True)

    # Tab 3: Hourly Analysis
    with tab3:
        st.header("Hourly Performance Breakdown")
        st.plotly_chart(create_hourly_analysis(data), use_container_width=True)

    # Tab 4: Time Waste
    with tab4:
        st.header("Time Waste Analysis")
        waiting_data = data[data['Category'] == 'Waiting']
        if not waiting_data.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Waiting Time Breakdown")
                fig = px.bar(
                    waiting_data.groupby('Description')['Duration'].sum().sort_values(ascending=False).head(10),
                    orientation='h', title="Top Time Waste Causes"
                )
                fig.update_traces(marker_color="#DC143C")
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                st.subheader("Waiting Time Impact")
                total_waiting = waiting_data['Duration'].sum()
                total_time = data['Duration'].sum()
                days_count = data['Sheet'].nunique()
                avg_waiting_per_day = (total_waiting / days_count) if days_count > 0 else 0
                st.metric("Total Waiting Time (all days)", f"{total_waiting:.0f} min")
                st.metric("Percentage of Total (all days)", f"{(total_waiting/total_time)*100:.1f}%")
                st.info(f"Insight: Eliminating waiting time could improve wrench time by {total_waiting:.0f} minutes across {days_count} day(s) (~{avg_waiting_per_day:.0f} min/day).")
        else:
            st.success("No waiting time detected.")

    # Tab 5: Incidental Activities
    with tab5:
        st.header("Incidental Activities Analysis")
        if not data[data['Category'] == 'Incidental'].empty:
            st.plotly_chart(create_incidental_analysis(data), use_container_width=True)
        else:
            st.info("No incidental activities found.")

    # Tab 6: Trends (only for multi-sheet)
    if data['Sheet'].nunique() > 1:
        with tab6:
            st.header("Trend Analysis")
            fig = px.line(
                data.groupby(['Sheet', 'Category'])['Duration'].sum().reset_index(),
                x='Sheet', y='Duration', color='Category', title="Daily Category Trends"
            )
            st.plotly_chart(fig, use_container_width=True)
    
    # Tab 7: Daily Analysis
    with tab7:
        st.header("Daily Analysis")
        sheets = data['Sheet'].unique()
        selected_sheet = st.selectbox("Select a Day/Sheet to Analyze", sheets)
        if selected_sheet:
            st.subheader(f"Analysis for: {selected_sheet}")
            sheet_data = data[data['Sheet'] == selected_sheet]
            col1, col2 = st.columns(2)
            with col1:
                total_daily = sheet_data['Duration'].sum()
                wrench_daily = sheet_data[sheet_data['Category'] == 'Wrench Time']['Duration'].sum()
                efficiency_daily = (wrench_daily / total_daily) * 100 if total_daily > 0 else 0
                st.metric("Total Time", f"{total_daily:.0f} min")
                st.metric("Wrench Time", f"{wrench_daily:.0f} min")
                st.metric("Daily Efficiency", f"{efficiency_daily:.1f}%")
                top_activities = sheet_data.groupby('Description')['Duration'].sum().sort_values(ascending=False).head(5)
                st.subheader("Top Activities")
                for activity, duration in top_activities.items():
                    st.write(f"• {activity}: {duration:.0f} min")
            with col2:
                st.plotly_chart(
                    create_donut_chart(sheet_data, f"{selected_sheet} - Category Distribution", height=300),
                    use_container_width=True,
                    key=f"chart_daily_donut_{selected_sheet}"
                )
            st.markdown("---")
            st.subheader("Activity Timeline")
            st.plotly_chart(
                create_activity_timeline(sheet_data),
                use_container_width=True,
                key=f"chart_activity_timeline_{selected_sheet}"
            )

if __name__ == "__main__":
    main()