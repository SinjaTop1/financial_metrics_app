from flask import Flask, render_template, request, send_file, jsonify
import yfinance as yf
import pandas as pd
import numpy as np
import json
import io
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import plotly
import plotly.express as px
import plotly.graph_objects as go

app = Flask(__name__)

# Define sample companies
companies = {
    "AAPL": "Apple Inc.",
    "MSFT": "Microsoft Corporation",
    "GOOGL": "Alphabet Inc.",
    "AMZN": "Amazon.com, Inc.",
    "META": "Meta Platforms, Inc.",
    "TSLA": "Tesla, Inc.",
    "NVDA": "NVIDIA Corporation",
    "JPM": "JPMorgan Chase & Co.",
    "WMT": "Walmart Inc.",
    "KO": "The Coca-Cola Company"
}

@app.route('/')
def index():
    return render_template('index.html', companies=companies)

# Helper function to find the closest matching row
def find_row(df, possible_names):
    for name in possible_names:
        if name in df.index:
            return df.loc[name]
    return None

# Function to calculate financial metrics
def calculate_metrics(ticker, years_back=5):
    try:
        # Get financial data
        company = yf.Ticker(ticker)
        
        # Get balance sheet, income statement, and cash flow data
        balance_sheet = company.balance_sheet
        income_stmt = company.income_stmt
        cash_flow = company.cashflow
        
        # Convert column names from timestamps to string dates for better display
        balance_sheet.columns = [col.strftime('%Y-%m-%d') if hasattr(col, 'strftime') else col for col in balance_sheet.columns]
        income_stmt.columns = [col.strftime('%Y-%m-%d') if hasattr(col, 'strftime') else col for col in income_stmt.columns]
        cash_flow.columns = [col.strftime('%Y-%m-%d') if hasattr(col, 'strftime') else col for col in cash_flow.columns]
        
        # Get common dates across all statements
        common_dates = sorted(set(balance_sheet.columns).intersection(set(income_stmt.columns)), reverse=True)
        
        # Limit to requested years
        common_dates = common_dates[:min(years_back, len(common_dates))]
        
        # Filter to common dates
        balance_sheet = balance_sheet[common_dates]
        income_stmt = income_stmt[common_dates]
        cash_flow = cash_flow[common_dates]
        
        # Create a metrics dataframe
        metrics = pd.DataFrame(index=common_dates)
        
        # Find key financial rows with alternative names
        # Current Assets
        current_assets = find_row(balance_sheet, [
            'Total Current Assets', 
            'CurrentAssets', 
            'Current Assets',
            'Total Current Assets'
        ])
        if current_assets is None:
            # Debug print balance sheet rows
            return None, None, None, None, f"Could not find Current Assets in balance sheet. Available rows: {', '.join(balance_sheet.index)}"
        
        # Current Liabilities
        current_liabilities = find_row(balance_sheet, [
            'Total Current Liabilities', 
            'CurrentLiabilities', 
            'Current Liabilities'
        ])
        if current_liabilities is None:
            return None, None, None, None, f"Could not find Current Liabilities in balance sheet. Available rows: {', '.join(balance_sheet.index)}"
        
        # Revenue
        revenue = find_row(income_stmt, [
            'Total Revenue', 
            'Revenue', 
            'TotalRevenue',
            'Gross Revenue',
            'Sales'
        ])
        if revenue is None:
            return None, None, None, None, f"Could not find Revenue in income statement. Available rows: {', '.join(income_stmt.index)}"
        
        # Net Income
        net_income = find_row(income_stmt, [
            'Net Income', 
            'NetIncome', 
            'Net Income Common Stockholders',
            'Net Income From Continuing Operations'
        ])
        if net_income is None:
            return None, None, None, None, f"Could not find Net Income in income statement. Available rows: {', '.join(income_stmt.index)}"
        
        # Total Assets
        total_assets = find_row(balance_sheet, [
            'Total Assets', 
            'TotalAssets', 
            'Assets'
        ])
        if total_assets is None:
            return None, None, None, None, f"Could not find Total Assets in balance sheet. Available rows: {', '.join(balance_sheet.index)}"
        
        # Stockholder Equity
        shareholder_equity = find_row(balance_sheet, [
            'Total Stockholder Equity', 
            'StockholderEquity', 
            'Stockholders Equity',
            'Total Shareholders Equity',
            'Shareholders Equity'
        ])
        if shareholder_equity is None:
            return None, None, None, None, f"Could not find Stockholder Equity in balance sheet. Available rows: {', '.join(balance_sheet.index)}"
        
        # Inventory (Optional)
        inventory = find_row(balance_sheet, [
            'Inventory', 
            'Inventories', 
            'Total Inventory'
        ])
        
        # Accounts Receivable (Optional)
        accounts_receivable = find_row(balance_sheet, [
            'Net Receivables', 
            'Accounts Receivable', 
            'AccountsReceivable',
            'Total Receivables'
        ])
        
        # Calculate metrics
        # Current Ratio = Current Assets / Current Liabilities
        metrics['Current Ratio'] = current_assets / current_liabilities
        
        # Quick Ratio = (Current Assets - Inventory) / Current Liabilities
        if inventory is not None:
            metrics['Quick Ratio'] = (current_assets - inventory) / current_liabilities
        else:
            # If inventory not available, use Current Ratio as approximation
            metrics['Quick Ratio'] = metrics['Current Ratio']
        
        # Current Asset Turnover = Revenue / Average Current Assets
        avg_current_assets = current_assets.copy()
        for i in range(1, len(current_assets)):
            avg_current_assets.iloc[i-1] = (current_assets.iloc[i-1] + current_assets.iloc[i]) / 2
        
        metrics['Current Asset Turnover'] = revenue / avg_current_assets
        
        # Total Asset Turnover = Revenue / Average Total Assets
        avg_total_assets = total_assets.copy()
        for i in range(1, len(total_assets)):
            avg_total_assets.iloc[i-1] = (total_assets.iloc[i-1] + total_assets.iloc[i]) / 2
        
        metrics['Total Asset Turnover'] = revenue / avg_total_assets
        
        # Days Sales Outstanding = (Accounts Receivable / Revenue) * 365
        if accounts_receivable is not None:
            metrics['Days Sales Outstanding'] = (accounts_receivable / (revenue / 365))
        else:
            metrics['Days Sales Outstanding'] = np.nan
        
        # Profit Margin = Net Income / Revenue
        metrics['Profit Margin'] = net_income / revenue
        
        # Total Liabilities (calculate if not directly available)
        total_liabilities = find_row(balance_sheet, [
            'Total Liabilities', 
            'TotalLiabilities',
            'Liabilities'
        ])
        if total_liabilities is None:
            # Calculate Total Liabilities by subtracting Total Stockholder Equity from Total Assets
            total_liabilities = total_assets - shareholder_equity
        
        # Debt Ratio = Total Liabilities / Total Assets
        metrics['Debt Ratio'] = total_liabilities / total_assets
        
        # Return on Equity = Net Income / Shareholders' Equity
        metrics['Return on Equity'] = net_income / shareholder_equity
        
        # EBIT for Basic Earning Power
        ebit = find_row(income_stmt, [
            'EBIT', 
            'Operating Income', 
            'OperatingIncome',
            'Income Before Tax'
        ])
        if ebit is None:
            # Calculate EBIT as Net Income + Interest Expense + Income Tax Expense
            ebit = net_income
            
            interest_expense = find_row(income_stmt, [
                'Interest Expense', 
                'InterestExpense'
            ])
            if interest_expense is not None:
                ebit += interest_expense
                
            income_tax = find_row(income_stmt, [
                'Income Tax Expense', 
                'IncomeTaxExpense',
                'Tax Provision',
                'Provision for Income Taxes'
            ])
            if income_tax is not None:
                ebit += income_tax
        
        # Basic Earning Power = EBIT / Total Assets
        metrics['Basic Earning Power'] = ebit / total_assets
        
        # Transpose the metrics for better display
        metrics_display = metrics.transpose()
        
        # Create chart data
        chart_data = {}
        
        # Liquidity Ratios Chart
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Current Ratio'].tolist(), mode='lines+markers', name='Current Ratio'))
        fig1.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Quick Ratio'].tolist(), mode='lines+markers', name='Quick Ratio'))
        fig1.update_layout(title='Liquidity Ratios', xaxis_title='Year', yaxis_title='Ratio Value')
        chart_data['liquidity'] = json.dumps(fig1, cls=plotly.utils.PlotlyJSONEncoder)
        
        # Efficiency Ratios Chart
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Current Asset Turnover'].tolist(), mode='lines+markers', name='Current Asset Turnover'))
        fig2.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Total Asset Turnover'].tolist(), mode='lines+markers', name='Total Asset Turnover'))
        fig2.update_layout(title='Asset Turnover Ratios', xaxis_title='Year', yaxis_title='Turnover Ratio')
        chart_data['efficiency'] = json.dumps(fig2, cls=plotly.utils.PlotlyJSONEncoder)
        
        # Profitability Ratios Chart
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Profit Margin'].tolist(), mode='lines+markers', name='Profit Margin'))
        fig3.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Return on Equity'].tolist(), mode='lines+markers', name='Return on Equity'))
        fig3.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Basic Earning Power'].tolist(), mode='lines+markers', name='Basic Earning Power'))
        fig3.update_layout(title='Profitability Ratios', xaxis_title='Year', yaxis_title='Ratio Value')
        chart_data['profitability'] = json.dumps(fig3, cls=plotly.utils.PlotlyJSONEncoder)
        
        # Debt Ratio Chart
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=metrics.index.tolist(), y=metrics['Debt Ratio'].tolist(), name='Debt Ratio'))
        fig4.update_layout(title='Debt Ratio', xaxis_title='Year', yaxis_title='Ratio Value')
        chart_data['solvency'] = json.dumps(fig4, cls=plotly.utils.PlotlyJSONEncoder)
        
        # Days Sales Outstanding Chart (if data available)
        if not metrics['Days Sales Outstanding'].isnull().all():
            fig5 = go.Figure()
            fig5.add_trace(go.Scatter(x=metrics.index.tolist(), y=metrics['Days Sales Outstanding'].tolist(), mode='lines+markers', name='Days Sales Outstanding'))
            fig5.update_layout(title='Days Sales Outstanding', xaxis_title='Year', yaxis_title='Days')
            chart_data['dso'] = json.dumps(fig5, cls=plotly.utils.PlotlyJSONEncoder)
        
        # Return the financial data and calculated metrics
        return balance_sheet, income_stmt, cash_flow, metrics_display, chart_data
        
    except Exception as e:
        return None, None, None, None, str(e)

@app.route('/analyze', methods=['POST'])
def analyze():
    ticker = request.form.get('company')
    years = int(request.form.get('years', 5))
    
    try:
        balance_sheet, income_stmt, cash_flow, metrics, chart_data = calculate_metrics(ticker, years)
        
        if metrics is None:
            return jsonify({'error': chart_data})  # chart_data contains error message
        
        # Convert DataFrames to HTML
        metrics_html = metrics.to_html(classes='data-table', float_format=lambda x: f'{x:.2f}')
        balance_sheet_html = balance_sheet.to_html(classes='data-table')
        income_stmt_html = income_stmt.to_html(classes='data-table')
        cash_flow_html = cash_flow.to_html(classes='data-table')
        
        return jsonify({
            'metrics': metrics_html,
            'balance_sheet': balance_sheet_html,
            'income_stmt': income_stmt_html,
            'cash_flow': cash_flow_html,
            'charts': chart_data
        })
    
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/download', methods=['POST'])
def download():
    ticker = request.form.get('company')
    years = int(request.form.get('years', 5))
    
    try:
        balance_sheet, income_stmt, cash_flow, metrics, chart_data = calculate_metrics(ticker, years)
        
        if metrics is None:
            return jsonify({'error': 'Failed to calculate metrics: ' + chart_data})
        
        # Create Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            metrics.to_excel(writer, sheet_name='Financial Metrics')
            balance_sheet.to_excel(writer, sheet_name='Balance Sheet')
            income_stmt.to_excel(writer, sheet_name='Income Statement')
            cash_flow.to_excel(writer, sheet_name='Cash Flow')
            
            # Add formulas sheet
            formulas = pd.DataFrame([
                ["Current Ratio", "Current Assets / Current Liabilities"],
                ["Quick Ratio", "(Current Assets - Inventory) / Current Liabilities"],
                ["Current Asset Turnover", "Revenue / Average Current Assets"],
                ["Total Asset Turnover", "Revenue / Average Total Assets"],
                ["Days Sales Outstanding", "(Accounts Receivable / Revenue) * 365"],
                ["Profit Margin", "Net Income / Revenue"],
                ["Debt Ratio", "Total Liabilities / Total Assets"],
                ["Return on Equity", "Net Income / Shareholders' Equity"],
                ["Basic Earning Power", "EBIT / Total Assets"]
            ], columns=["Metric", "Formula"])
            formulas.to_excel(writer, sheet_name='Formulas', index=False)
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{ticker}_financial_metrics.xlsx"
        )
    
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)