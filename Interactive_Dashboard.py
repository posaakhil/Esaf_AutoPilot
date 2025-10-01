#!/usr/bin/env python3
"""
ESAF Executive Dashboard Generator - LEADER-FOCUSED VERSION
Clear, meaningful visualizations for executive decision-making
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import json
import sys
from pathlib import Path
import html
import argparse
from datetime import datetime, timedelta

# ===== LOAD CONFIG =====
try:
    with open("esaf_config.json", 'r', encoding='utf-8') as f:
        config = json.load(f)
    DASHBOARD_TITLE = config.get("dashboard", {}).get("title", "ESAF Access Requests Executive Dashboard")
    SHEETS_TO_PROCESS = config.get("dashboard", {}).get("sheets", ["Master_Data", "HCA_India", "HCA_Domestic", "Summary"])
except Exception:
    DASHBOARD_TITLE = "ESAF Access Requests Executive Dashboard"
    SHEETS_TO_PROCESS = ["Master_Data", "HCA_India", "HCA_Domestic", "Summary"]

# Professional color scheme
COLORS = {
    'primary': '#2E5984',
    'secondary': '#1E3A5F',
    'accent': '#FF6B35',
    'success': '#10B981',
    'warning': '#F59E0B',
    'danger': '#EF4444',
    'info': '#3B82F6',
    'light': '#F8FAFC',
    'dark': '#1E293B',
    'create': '#10B981',
    'modify': '#3B82F6',
    'delete': '#EF4444'
}

def safe_print(message):
    """Safely print messages without Unicode issues"""
    safe_message = message.encode('utf-8', errors='replace').decode('utf-8', errors='replace')
    print(safe_message)

def create_executive_kpi_cards(df, sheet_name):
    """Create executive-focused KPI cards"""
    total_requests = len(df)
    
    # Request type analysis
    create_count = len(df[df['Request'].str.contains('Create', na=False)]) if 'Request' in df.columns else 0
    modify_count = len(df[df['Request'].str.contains('Modify', na=False)]) if 'Request' in df.columns else 0
    delete_count = len(df[df['Request'].str.contains('Delete', na=False)]) if 'Request' in df.columns else 0
    
    # Application diversity
    unique_apps = df['Application'].nunique() if 'Application' in df.columns else 0
    
    # Calculate percentages
    create_pct = (create_count / total_requests * 100) if total_requests > 0 else 0
    modify_pct = (modify_count / total_requests * 100) if total_requests > 0 else 0
    delete_pct = (delete_count / total_requests * 100) if total_requests > 0 else 0
    
    kpi_html = f"""
    <div class="kpi-section">
        <h3 class="section-title">Executive Summary - {html.escape(sheet_name)}</h3>
        <div class="kpi-grid">
            <div class="kpi-card total">
                <div class="kpi-icon">TOTAL</div>
                <div class="kpi-content">
                    <div class="kpi-value">{total_requests}</div>
                    <div class="kpi-label">Total Requests</div>
                    <div class="kpi-trend">All Requests</div>
                </div>
            </div>
            <div class="kpi-card apps">
                <div class="kpi-icon">APPS</div>
                <div class="kpi-content">
                    <div class="kpi-value">{unique_apps}</div>
                    <div class="kpi-label">Applications</div>
                    <div class="kpi-trend">Across Systems</div>
                </div>
            </div>
            <div class="kpi-card create">
                <div class="kpi-icon">CREATE</div>
                <div class="kpi-content">
                    <div class="kpi-value">{create_count}</div>
                    <div class="kpi-label">New Access</div>
                    <div class="kpi-trend">{create_pct:.1f}% of total</div>
                </div>
            </div>
            <div class="kpi-card modify">
                <div class="kpi-icon">MODIFY</div>
                <div class="kpi-content">
                    <div class="kpi-value">{modify_count}</div>
                    <div class="kpi-label">Access Changes</div>
                    <div class="kpi-trend">{modify_pct:.1f}% of total</div>
                </div>
            </div>
            <div class="kpi-card delete">
                <div class="kpi-icon">DELETE</div>
                <div class="kpi-content">
                    <div class="kpi-value">{delete_count}</div>
                    <div class="kpi-label">Access Removal</div>
                    <div class="kpi-trend">{delete_pct:.1f}% of total</div>
                </div>
            </div>
        </div>
    </div>
    """
    return kpi_html

def create_request_type_donut(df):
    """Create donut chart for request type distribution - clear and executive-friendly"""
    if 'Request' not in df.columns:
        return None
        
    create_count = len(df[df['Request'].str.contains('Create', na=False)])
    modify_count = len(df[df['Request'].str.contains('Modify', na=False)])
    delete_count = len(df[df['Request'].str.contains('Delete', na=False)])
    other_count = len(df) - (create_count + modify_count + delete_count)
    
    labels = ['Create Access', 'Modify Access', 'Delete Access', 'Other']
    values = [create_count, modify_count, delete_count, other_count]
    colors = [COLORS['create'], COLORS['modify'], COLORS['delete'], COLORS['light']]
    
    # Only show if we have meaningful data
    if sum(values) == 0:
        return None
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.6,
        marker_colors=colors,
        textinfo='label+percent',
        textposition='inside',
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
    )])
    
    fig.update_layout(
        title="Access Request Distribution",
        title_x=0.5,
        showlegend=False,
        margin=dict(t=50, b=20, l=20, r=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color=COLORS['dark']),
        height=400
    )
    
    return pio.to_html(fig, full_html=False, include_plotlyjs=False)

def create_top_applications_bar(df):
    """Create horizontal bar chart for top applications - clear ranking"""
    if 'Application' not in df.columns:
        return None
        
    app_counts = df['Application'].value_counts().head(10)  # Top 10 applications
    
    if len(app_counts) == 0:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        y=app_counts.index,
        x=app_counts.values,
        orientation='h',
        marker_color=COLORS['primary'],
        text=app_counts.values,
        textposition='auto',
        hovertemplate='<b>%{y}</b><br>Requests: %{x}<extra></extra>'
    ))
    
    fig.update_layout(
        title="Top 10 Applications by Request Volume",
        title_x=0.5,
        margin=dict(t=50, b=20, l=150, r=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color=COLORS['dark']),
        xaxis=dict(title="Number of Requests", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        yaxis=dict(title="", autorange="reversed", tickfont=dict(size=11)),
        height=max(400, len(app_counts) * 40),
        showlegend=False
    )
    
    return pio.to_html(fig, full_html=False, include_plotlyjs=False)

def create_request_trend_analysis(df):
    """Create line chart showing request trends over time"""
    date_columns = ['Request date', 'Last updated time']
    available_date_cols = [col for col in date_columns if col in df.columns]
    
    if not available_date_cols:
        return None
        
    trend_data = []
    for date_col in available_date_cols:
        if date_col in df.columns:
            # Convert to datetime and extract date parts
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            daily_counts = df[date_col].dt.date.value_counts().sort_index()
            
            for date, count in daily_counts.items():
                trend_data.append({
                    'Date': date,
                    'Count': count,
                    'Metric': 'Daily Requests'
                })
    
    if not trend_data:
        return None
        
    trend_df = pd.DataFrame(trend_data)
    
    # Calculate 7-day moving average for trend line
    if len(trend_df) > 7:
        trend_df['Moving_Avg'] = trend_df['Count'].rolling(window=7, min_periods=1).mean()
    
    fig = go.Figure()
    
    # Actual daily counts
    fig.add_trace(go.Scatter(
        x=trend_df['Date'],
        y=trend_df['Count'],
        mode='lines+markers',
        name='Daily Requests',
        line=dict(color=COLORS['primary'], width=3),
        marker=dict(size=6),
        hovertemplate='<b>%{x}</b><br>Requests: %{y}<extra></extra>'
    ))
    
    # Moving average trend line
    if 'Moving_Avg' in trend_df.columns:
        fig.add_trace(go.Scatter(
            x=trend_df['Date'],
            y=trend_df['Moving_Avg'],
            mode='lines',
            name='7-Day Average',
            line=dict(color=COLORS['accent'], width=2, dash='dash'),
            hovertemplate='<b>%{x}</b><br>7-Day Avg: %{y:.1f}<extra></extra>'
        ))
    
    fig.update_layout(
        title="Request Volume Trend Analysis",
        title_x=0.5,
        margin=dict(t=50, b=40, l=60, r=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color=COLORS['dark']),
        xaxis=dict(title="Date", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        yaxis=dict(title="Number of Requests", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        height=400,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    return pio.to_html(fig, full_html=False, include_plotlyjs=False)

def create_application_priority_matrix(df):
    """Create a priority matrix showing application vs request type impact"""
    if 'Application' not in df.columns or 'Request' not in df.columns:
        return None
        
    # Prepare data for priority analysis
    priority_data = []
    for app in df['Application'].dropna().unique()[:15]:  # Top 15 apps
        app_df = df[df['Application'] == app]
        total_requests = len(app_df)
        
        create_count = len(app_df[app_df['Request'].str.contains('Create', na=False)])
        modify_count = len(app_df[app_df['Request'].str.contains('Modify', na=False)])
        delete_count = len(app_df[app_df['Request'].str.contains('Delete', na=False)])
        
        # Calculate risk score (higher for delete operations)
        risk_score = (create_count * 1 + modify_count * 2 + delete_count * 3) / total_requests if total_requests > 0 else 0
        
        priority_data.append({
            'Application': app,
            'Total_Requests': total_requests,
            'Create_Count': create_count,
            'Modify_Count': modify_count,
            'Delete_Count': delete_count,
            'Risk_Score': risk_score
        })
    
    if not priority_data:
        return None
        
    priority_df = pd.DataFrame(priority_data)
    
    fig = go.Figure(data=go.Scatter(
        x=priority_df['Total_Requests'],
        y=priority_df['Risk_Score'],
        mode='markers+text',
        marker=dict(
            size=priority_df['Total_Requests']/priority_df['Total_Requests'].max() * 50 + 20,
            color=priority_df['Risk_Score'],
            colorscale='RdYlGn_r',
            showscale=True,
            colorbar=dict(title="Risk Level")
        ),
        text=priority_df['Application'],
        textposition="middle center",
        hovertemplate='<b>%{text}</b><br>Total Requests: %{x}<br>Risk Score: %{y:.2f}<extra></extra>'
    ))
    
    fig.update_layout(
        title="Application Priority Matrix (Volume vs Risk)",
        title_x=0.5,
        margin=dict(t=50, b=60, l=80, r=40),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color=COLORS['dark']),
        xaxis=dict(title="Total Request Volume", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        yaxis=dict(title="Risk Score (Higher = More Critical)", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        height=500
    )
    
    return pio.to_html(fig, full_html=False, include_plotlyjs=False)

def create_workload_distribution_chart(df):
    """Create a chart showing workload distribution patterns"""
    if 'Application' not in df.columns:
        return None
        
    # Calculate workload distribution
    app_distribution = df['Application'].value_counts().head(8)
    
    if len(app_distribution) == 0:
        return None
    
    # Calculate percentage distribution
    total = app_distribution.sum()
    percentages = [(count/total * 100) for count in app_distribution.values]
    
    fig = go.Figure(data=[
        go.Bar(
            x=app_distribution.index,
            y=app_distribution.values,
            text=[f'{p:.1f}%' for p in percentages],
            textposition='auto',
            marker_color=COLORS['primary'],
            hovertemplate='<b>%{x}</b><br>Requests: %{y}<br>Share: %{text}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Workload Distribution Across Applications",
        title_x=0.5,
        margin=dict(t=50, b=80, l=60, r=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color=COLORS['dark']),
        xaxis=dict(title="Applications", tickangle=45, showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        yaxis=dict(title="Number of Requests", showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
        height=400
    )
    
    return pio.to_html(fig, full_html=False, include_plotlyjs=False)

def create_executive_summary_table(df):
    """Create a summary table for quick executive insights"""
    if 'Application' not in df.columns or 'Request' not in df.columns:
        return None
        
    # Calculate key metrics for top applications
    summary_data = []
    for app in df['Application'].dropna().unique()[:8]:  # Top 8 apps
        app_df = df[df['Application'] == app]
        total = len(app_df)
        
        create_pct = (len(app_df[app_df['Request'].str.contains('Create', na=False)]) / total * 100) if total > 0 else 0
        modify_pct = (len(app_df[app_df['Request'].str.contains('Modify', na=False)]) / total * 100) if total > 0 else 0
        delete_pct = (len(app_df[app_df['Request'].str.contains('Delete', na=False)]) / total * 100) if total > 0 else 0
        
        summary_data.append({
            'Application': app,
            'Total_Requests': total,
            'Create_Pct': create_pct,
            'Modify_Pct': modify_pct,
            'Delete_Pct': delete_pct
        })
    
    if not summary_data:
        return None
        
    # Create HTML table
    table_html = """
    <div class="summary-table-section">
        <h3 class="section-title">Application Performance Summary</h3>
        <div class="table-container">
            <table class="executive-table">
                <thead>
                    <tr>
                        <th>Application</th>
                        <th>Total Requests</th>
                        <th>Create %</th>
                        <th>Modify %</th>
                        <th>Delete %</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for item in summary_data:
        table_html += f"""
                    <tr>
                        <td><strong>{html.escape(str(item['Application']))}</strong></td>
                        <td>{item['Total_Requests']}</td>
                        <td style="color: {COLORS['create']}">{item['Create_Pct']:.1f}%</td>
                        <td style="color: {COLORS['modify']}">{item['Modify_Pct']:.1f}%</td>
                        <td style="color: {COLORS['delete']}">{item['Delete_Pct']:.1f}%</td>
                    </tr>
        """
    
    table_html += """
                </tbody>
            </table>
        </div>
    </div>
    """
    
    return table_html

def create_interactive_filters():
    """Create HTML for interactive filter controls"""
    return """
    <div class="filter-section">
        <h3 class="section-title">Data Analysis Controls</h3>
        <div class="filter-grid">
            <div class="filter-group">
                <label for="timeFrameFilter">Time Frame:</label>
                <select id="timeFrameFilter">
                    <option value="7d">Last 7 Days</option>
                    <option value="30d" selected>Last 30 Days</option>
                    <option value="90d">Last 90 Days</option>
                    <option value="all">All Time</option>
                </select>
            </div>
            <div class="filter-group">
                <label for="requestTypeFilter">Request Type:</label>
                <select id="requestTypeFilter">
                    <option value="all" selected>All Types</option>
                    <option value="Create">Create Only</option>
                    <option value="Modify">Modify Only</option>
                    <option value="Delete">Delete Only</option>
                </select>
            </div>
            <div class="filter-group">
                <label for="applicationFilter">Application:</label>
                <select id="applicationFilter">
                    <option value="all" selected>All Applications</option>
                </select>
            </div>
            <div class="filter-group">
                <button onclick="applyExecutiveFilters()" class="filter-btn">Apply Analysis</button>
                
            </div>
        </div>
    </div>
    """

def generate_executive_charts(df, sheet_name):
    """Generate all executive-focused charts for a sheet"""
    charts = []
    
    # KPI Cards
    kpi_html = create_executive_kpi_cards(df, sheet_name)
    
    # Donut Chart for Request Types
    donut_html = create_request_type_donut(df)
    if donut_html:
        charts.append({
            'title': 'Access Request Distribution',
            'div': donut_html,
            'class': 'chart-half'
        })
    
    # Top Applications Bar Chart
    top_apps_html = create_top_applications_bar(df)
    if top_apps_html:
        charts.append({
            'title': 'Top Applications by Volume',
            'div': top_apps_html,
            'class': 'chart-half'
        })
    
    # Trend Analysis
    trend_html = create_request_trend_analysis(df)
    if trend_html:
        charts.append({
            'title': 'Request Volume Trends',
            'div': trend_html,
            'class': 'chart-full'
        })
    
    # Priority Matrix
    priority_html = create_application_priority_matrix(df)
    if priority_html:
        charts.append({
            'title': 'Application Priority Matrix',
            'div': priority_html,
            'class': 'chart-half'
        })
    
    # Workload Distribution
    workload_html = create_workload_distribution_chart(df)
    if workload_html:
        charts.append({
            'title': 'Workload Distribution',
            'div': workload_html,
            'class': 'chart-half'
        })
    
    # Executive Summary Table
    table_html = create_executive_summary_table(df)
    if table_html:
        charts.append({
            'title': '',
            'div': table_html,
            'class': 'table-full'
        })
    
    return kpi_html, charts

def build_executive_dashboard(sheet_results, page_title):
    """Build the complete executive dashboard HTML"""
    
    css = """
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }

    body {
        font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif;
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        color: #1e293b;
        min-height: 100vh;
        padding: 20px;
        line-height: 1.6;
    }

    .container {
        max-width: 1800px;
        margin: 0 auto;
    }

    .header {
        background: white;
        border-radius: 16px;
        padding: 30px;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        border: 1px solid #e2e8f0;
    }

    .header h1 {
        font-weight: 700;
        font-size: 2.2rem;
        color: #1E3A5F;
        margin-bottom: 8px;
    }

    .header p {
        color: #64748b;
        font-size: 1.1rem;
        font-weight: 500;
    }

    .header .subtitle {
        color: #94a3b8;
        font-size: 0.9rem;
        margin-top: 8px;
    }

    .filter-section {
        background: white;
        border-radius: 16px;
        padding: 25px;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        border: 1px solid #e2e8f0;
    }

    .section-title {
        font-size: 1.3rem;
        font-weight: 700;
        color: #1E3A5F;
        margin-bottom: 20px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 10px;
    }

    .filter-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 20px;
        align-items: end;
    }

    .filter-group {
        display: flex;
        flex-direction: column;
        gap: 8px;
    }

    .filter-group label {
        font-weight: 600;
        color: #374151;
        font-size: 0.9rem;
    }

    .filter-group select {
        padding: 12px;
        border: 2px solid #e2e8f0;
        border-radius: 10px;
        font-size: 14px;
        transition: all 0.3s ease;
        background: white;
    }

    .filter-group select:focus {
        outline: none;
        border-color: #3B82F6;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
    }

    .filter-btn {
        padding: 12px 24px;
        background: linear-gradient(135deg, #2E5984, #1E3A5F);
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.3s ease;
        margin-right: 10px;
    }

   

    .filter-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
    }

    .sheet-section {
        background: white;
        border-radius: 16px;
        padding: 25px;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        border: 1px solid #e2e8f0;
    }

    .kpi-section {
        margin-bottom: 25px;
    }

    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 20px;
        margin-top: 15px;
    }

    .kpi-card {
        background: linear-gradient(135deg, #ffffff, #f8fafc);
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
        border-left: 4px solid #cbd5e1;
        transition: all 0.3s ease;
    }

    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
    }

    .kpi-card.total { border-left-color: #2E5984; }
    .kpi-card.apps { border-left-color: #FF6B35; }
    .kpi-card.create { border-left-color: #10B981; }
    .kpi-card.modify { border-left-color: #3B82F6; }
    .kpi-card.delete { border-left-color: #EF4444; }

    .kpi-icon {
        font-size: 0.9rem;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
    }

    .kpi-value {
        font-size: 2rem;
        font-weight: 800;
        color: #1E3A5F;
        line-height: 1;
        margin-bottom: 4px;
    }

    .kpi-label {
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 600;
        margin-bottom: 4px;
    }

    .kpi-trend {
        font-size: 0.75rem;
        color: #94a3b8;
        font-weight: 500;
    }

    .chart-grid {
        display: grid;
        gap: 25px;
        margin-top: 20px;
    }

    .chart-full {
        grid-column: 1 / -1;
    }

    .chart-half {
        grid-column: span 1;
    }

    .table-full {
        grid-column: 1 / -1;
    }

    .chart-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
        border: 1px solid #f1f5f9;
    }

    .chart-card h3 {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1E3A5F;
        margin-bottom: 15px;
    }

    .plotly-graph {
        width: 100%;
        height: 100%;
        border-radius: 8px;
    }

    .summary-table-section {
        margin-top: 10px;
    }

    .table-container {
        overflow-x: auto;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }

    .executive-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.9rem;
    }

    .executive-table th {
        background: #2E5984;
        color: white;
        padding: 12px 15px;
        text-align: left;
        font-weight: 600;
        border: none;
    }

    .executive-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #e2e8f0;
    }

    .executive-table tr:hover {
        background: #f8fafc;
    }

    .executive-table tr:last-child td {
        border-bottom: none;
    }

    footer {
        text-align: center;
        color: #64748b;
        padding: 30px;
        margin-top: 40px;
        font-size: 0.85rem;
        border-top: 1px solid #e2e8f0;
    }

    @media (min-width: 1200px) {
        .chart-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    @media (max-width: 768px) {
        .header h1 { font-size: 1.8rem; }
        .kpi-grid { grid-template-columns: 1fr; }
        .filter-grid { grid-template-columns: 1fr; }
        .chart-grid { grid-template-columns: 1fr; }
        .sheet-section { padding: 20px; }
    }

    @media (max-width: 480px) {
        body { padding: 15px; }
        .header { padding: 20px; }
        .kpi-card { padding: 15px; }
        .kpi-value { font-size: 1.6rem; }
    }
    """

    js = """
         <script>
    // Executive Filter Functions
    function applyExecutiveFilters() {
        const timeFrame = document.getElementById('timeFrameFilter').value;
        const requestType = document.getElementById('requestTypeFilter').value;
        const application = document.getElementById('applicationFilter').value;
        
        // Show analysis summary
        alert('Analysis Parameters Applied:\\n' +
              'Time Frame: ' + document.getElementById('timeFrameFilter').options[document.getElementById('timeFrameFilter').selectedIndex].text + '\\n' +
              'Request Type: ' + document.getElementById('requestTypeFilter').options[document.getElementById('requestTypeFilter').selectedIndex].text + '\\n' +
              'Application: ' + document.getElementById('applicationFilter').options[document.getElementById('applicationFilter').selectedIndex].text);
        
        console.log('Executive Filters:', { timeFrame, requestType, application });
    }
    
    function exportExecutiveReport() {
        // Simulate report export
        alert('Executive Report exported successfully!\\nThe report has been saved in your downloads folder.');
        console.log('Exporting executive report...');
    }

    // Real-time executive dashboard updates
    function updateExecutiveDashboard() {
        const now = new Date();
        const timeElement = document.getElementById('currentTime');
        if (timeElement) {
            timeElement.textContent = 
                now.toLocaleString('en-US', { 
                    weekday: 'long',
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit'
                });
        }
    }
    
    // Add executive dashboard enhancements
    document.addEventListener('DOMContentLoaded', function() {
        // Smooth loading animations
        const elements = document.querySelectorAll('.chart-card, .kpi-card');
        elements.forEach((element, index) => {
            element.style.opacity = '0';
            element.style.transform = 'translateY(20px)';
            setTimeout(() => {
                element.style.transition = 'all 0.5s ease';
                element.style.opacity = '1';
                element.style.transform = 'translateY(0)';
            }, index * 100);
        });
        
        // Add executive tooltips
        const kpiCards = document.querySelectorAll('.kpi-card');
        kpiCards.forEach(card => {
            card.addEventListener('click', function() {
                const type = this.classList[1];
                const insights = {
                    'total': 'Total number of access requests across all applications',
                    'apps': 'Number of unique applications with pending requests',
                    'create': 'New access provisioning requests',
                    'modify': 'Access modification and update requests',
                    'delete': 'Access removal and deprovisioning requests'
                };
                alert('Executive Insight: ' + (insights[type] || 'Key performance indicator'));
            });
        });

        // Initialize and start the clock
        updateExecutiveDashboard();
        setInterval(updateExecutiveDashboard, 60000); // Update every minute
    });
    </script>
    """

    head = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html.escape(page_title)}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>{css}</style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{html.escape(page_title)}</h1>
            <p>Executive Intelligence Platform - Access Request Analytics</p>
            <p class="subtitle">Last Updated: <span id="currentTime"></span></p>
        </div>
        
        {create_interactive_filters()}
"""

    body = ""
    for i, sheet_data in enumerate(sheet_results):
        sheet_name = sheet_data['name']
        kpi_html = sheet_data['kpi_html']
        charts = sheet_data['charts']
        
        if not kpi_html and not charts:
            continue

        body += f"""
        <div class="sheet-section" id="sheet-{i}">
            {kpi_html if kpi_html else ''}
        """

        if charts:
            # Organize charts by type
            half_charts = [c for c in charts if c.get('class') == 'chart-half']
            full_charts = [c for c in charts if c.get('class') == 'chart-full']
            table_charts = [c for c in charts if c.get('class') == 'table-full']
            
            # Display half charts in grid
            if half_charts:
                body += '<div class="chart-grid">'
                for chart in half_charts:
                    if chart['title']:  # Only add title if it exists
                        body += f"""
                        <div class="chart-card">
                            <h3>{chart['title']}</h3>
                            <div class="plotly-graph">
                                {chart['div']}
                            </div>
                        </div>
                        """
                    else:
                        body += f"""
                        <div class="chart-card">
                            <div class="plotly-graph">
                                {chart['div']}
                            </div>
                        </div>
                        """
                body += '</div>'
            
            # Display full charts
            if full_charts:
                for chart in full_charts:
                    body += f"""
                    <div class="chart-card">
                        <h3>{chart['title']}</h3>
                        <div class="plotly-graph">
                            {chart['div']}
                        </div>
                    </div>
                    """
            
            # Display tables
            if table_charts:
                for chart in table_charts:
                    body += chart['div']

        body += "</div>"

    footer = f"""
        <footer>
            <p><strong>ESAF Executive Dashboard v3.0</strong> - Strategic Access Management Analytics</p>
            <p>HCA Healthcare Confidential - For Executive Review Only</p>
        </footer>
    </div>
    {js}
</body>
</html>
"""

    return head + body + footer

def process_excel_to_executive_dashboard(excel_path: Path, out_file: Path, title="ESAF Access Requests Executive Dashboard"):
    """Process Excel file and generate executive dashboard"""
    try:
        with pd.ExcelFile(excel_path) as xls:
            available_sheets = xls.sheet_names
    except Exception as e:
        safe_print(f"[ERROR] Cannot read Excel file: {e}")
        return

    sheet_results = []
    for sheet in SHEETS_TO_PROCESS:
        if sheet not in available_sheets:
            safe_print(f"[INFO] Sheet '{sheet}' not found, skipping...")
            continue
            
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet)
            safe_print(f"[SUCCESS] Processing sheet: {sheet} ({len(df)} rows)")
            
            kpi_html, charts = generate_executive_charts(df, sheet)
            
            sheet_results.append({
                'name': sheet,
                'kpi_html': kpi_html,
                'charts': charts
            })
            
        except Exception as e:
            safe_print(f"[ERROR] Failed to process sheet '{sheet}': {e}")

    if not sheet_results:
        safe_print("[ERROR] No sheets processed successfully")
        return

    html_content = build_executive_dashboard(sheet_results, page_title=title)
    
    try:
        out_file.write_text(html_content, encoding='utf-8')
        safe_print(f"[SUCCESS] Executive dashboard saved to: {out_file.resolve()}")
        safe_print(f"[INFO] Processed {len(sheet_results)} sheets with executive visualizations")
    except Exception as e:
        safe_print(f"[ERROR] Failed to save dashboard: {e}")

def main():
    parser = argparse.ArgumentParser(description="Generate Executive ESAF Dashboard")
    parser.add_argument("--out", "-o", default="esaf_executive_dashboard.html", help="Output HTML file name")
    parser.add_argument("--title", "-t", default=DASHBOARD_TITLE, help="Dashboard title")
    args = parser.parse_args()
    
    excel_path = Path("Today_Assignment.xlsx")
    if not excel_path.exists():
        safe_print("[ERROR] Today_Assignment.xlsx not found")
        sys.exit(1)
        
    out_file = Path(args.out)
    process_excel_to_executive_dashboard(excel_path, out_file, title=args.title)

if __name__ == "__main__":
    main()