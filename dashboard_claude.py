import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, dash_table
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GenialDashboard:
    def __init__(self):
        self.data = None
        self.net_sector_summary = None
        self.daily_evolution = None
        self.evol_sector = None
        self.total_net_by_type = None
        self.volume_total_by_type = None
        
    def load_data(self):
        """Carrega e processa os dados do Excel"""
        try:
            # Busca o Excel mais recente na pasta Downloads
            download_folder = Path.home() / "Downloads"
            excel_files = list(download_folder.glob("Brokeragem*.xlsx"))
            
            if not excel_files:
                raise FileNotFoundError("Nenhum arquivo 'Brokeragem*.xlsx' encontrado na pasta Downloads.")
            
            latest_file = max(excel_files, key=os.path.getmtime)
            logger.info(f"Carregando arquivo: {latest_file}")
            
            # Leitura e limpeza dos dados
            self.data = pd.read_excel(latest_file, engine='openpyxl')
            self._clean_data()
            self._process_data()
            
            logger.info("Dados carregados e processados com sucesso!")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao carregar dados: {str(e)}")
            return False
    
    def _clean_data(self):
        """Limpa e padroniza os dados"""
        # Converte a data para index
        self.data.index = pd.to_datetime(self.data['DT_NEGOCIO'])
        self.data = self.data.drop(columns=['DT_NEGOCIO'])
        
        # Renomeia colunas
        column_mapping = {
            'CD_SETOR': 'SECTOR',
            'CD_SUBSETOR': 'SUB-SECTOR',
            'CD_SEGMENTO': 'SEGMENT',
            'CONTA': 'INV TYPE',
            'VL_COMPRA': 'BUY R$',
            'VL_VENDA': 'SELL R$',
            'VL_NET': 'NET R$',
            'VL_TOTAL': 'TOTAL'
        }
        self.data.rename(columns=column_mapping, inplace=True)
        
        # Preenche valores nulos
        self.data.fillna("Others", inplace=True)
        
        # Padroniza setores
        self.data['SECTOR'] = self.data['SECTOR'].replace({'FII': 'Others', 'IBOV': 'Others'})
        
        # Padroniza tipos de investidores
        investor_mapping = {
            'ESTRANGEIRO': 'FOREIGNERS',
            'LOCAL INSTITUCIONAL': 'LOCALS',
            'GENIAL': 'RETAIL'
        }
        self.data['INV TYPE'] = self.data['INV TYPE'].replace(investor_mapping)
        
        # Filtra apenas investidores válidos
        valid_investors = ['LOCALS', 'FOREIGNERS', 'RETAIL']
        self.data = self.data[self.data['INV TYPE'].isin(valid_investors)]
    
    def _process_data(self):
        """Processa os dados para análise"""
        # Agregações por setor
        self.net_sector_summary = self.data.groupby(['INV TYPE', 'SECTOR'])['NET R$'].sum().reset_index()
        self.net_sector_summary['NET R$'] /= 1_000_000
        
        # Net por tipo de investidor
        self.total_net_by_type = self.data.groupby('INV TYPE')['NET R$'].sum() / 1_000_000
        
        # Total operado por investidor
        self.volume_total_by_type = self.data.groupby('INV TYPE')['TOTAL'].sum() / 1_000_000
        
        # Evolução diária
        self.daily_evolution = self.data.pivot_table(
            values='NET R$', 
            index=self.data.index, 
            columns='INV TYPE', 
            aggfunc='sum'
        )
        self.daily_evolution.fillna(0, inplace=True)
        self.daily_evolution /= 1_000_000
        
        # Evolução por setor
        self.evol_sector = self.data.pivot_table(
            values='NET R$', 
            index=self.data.index, 
            columns=['INV TYPE', 'SECTOR'], 
            aggfunc='sum'
        )
        self.evol_sector.fillna(0, inplace=True)
        self.evol_sector /= 1_000_000
    
    def get_top_bottom_sectors(self, inv_type, n=3):
        """Retorna top e bottom setores por tipo de investidor"""
        filtered = self.net_sector_summary[
            self.net_sector_summary['INV TYPE'] == inv_type
        ].sort_values(by='NET R$', ascending=False)
        
        top_n = filtered.head(n)
        bottom_n = filtered.tail(n)
        
        return top_n, bottom_n
    
    def create_net_flow_chart(self):
        """Cria gráfico de fluxo líquido por setor"""
        fig = go.Figure()
        
        colors = {'LOCALS': '#1f77b4', 'FOREIGNERS': '#ff7f0e', 'RETAIL': '#2ca02c'}
        
        for inv in self.net_sector_summary['INV TYPE'].unique():
            data_inv = self.net_sector_summary[self.net_sector_summary['INV TYPE'] == inv]
            fig.add_trace(go.Bar(
                name=inv,
                x=data_inv['SECTOR'],
                y=data_inv['NET R$'],
                marker_color=colors.get(inv, '#d62728'),
                hovertemplate='<b>%{x}</b><br>%{y:.1f}M<extra></extra>'
            ))
        
        fig.update_layout(
            title="Net Flow by Investor Type and Sector",
            barmode='group',
            yaxis_title="R$ Million",
            plot_bgcolor='#f8f9fa',
            paper_bgcolor='#ffffff',
            font=dict(family="Arial", size=12),
            title_font=dict(size=16, color='#003366'),
            height=500
        )
        
        return fig
    
    def create_daily_evolution_chart(self):
        """Cria gráfico de evolução diária"""
        fig = go.Figure()
        
        colors = {'LOCALS': '#1f77b4', 'FOREIGNERS': '#ff7f0e', 'RETAIL': '#2ca02c'}
        
        for col in self.daily_evolution.columns:
            fig.add_trace(go.Scatter(
                x=self.daily_evolution.index,
                y=self.daily_evolution[col],
                name=col,
                mode='lines+markers',
                line=dict(color=colors.get(col, '#d62728'), width=2),
                hovertemplate='<b>%{fullData.name}</b><br>%{x}<br>%{y:.1f}M<extra></extra>'
            ))
        
        fig.update_layout(
            title="Daily Net Flow by Investor Type",
            yaxis_title="R$ Million",
            plot_bgcolor='#f8f9fa',
            paper_bgcolor='#ffffff',
            font=dict(family="Arial", size=12),
            title_font=dict(size=16, color='#003366'),
            height=400
        )
        
        return fig
    
    def create_volume_pie_chart(self):
        """Cria gráfico de pizza do volume"""
        fig = go.Figure(data=[go.Pie(
            labels=self.volume_total_by_type.index,
            values=self.volume_total_by_type.values,
            hole=0.4,
            hovertemplate='<b>%{label}</b><br>%{value:.1f}M (%{percent})<extra></extra>',
            textinfo='label+percent',
            textposition='outside'
        )])
        
        fig.update_layout(
            title="Share of Total Traded Volume",
            plot_bgcolor='#f8f9fa',
            paper_bgcolor='#ffffff',
            font=dict(family="Arial", size=12),
            title_font=dict(size=16, color='#003366'),
            height=400
        )
        
        return fig
    
    def create_app(self):
        """Cria a aplicação Dash"""
        app = Dash(__name__)
        app.title = "Genial Investor Flow"
        
        # Carrega os dados
        if not self.load_data():
            # Se falhar, cria uma página de erro
            app.layout = html.Div([
                html.H1("Erro ao carregar dados", style={'color': 'red', 'textAlign': 'center'}),
                html.P("Verifique se o arquivo Brokeragem*.xlsx está na pasta Downloads")
            ])
            return app
        
        # Cria o layout
        app.layout = html.Div([
            html.Div([
                html.H1("Genial - Investor Flow Dashboard", 
                       style={'textAlign': 'center', 'color': '#003366', 'marginBottom': '30px'}),
                html.P(f"Dados atualizados em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 
                      style={'textAlign': 'center', 'color': '#666', 'fontSize': '14px'})
            ], style={'marginBottom': '20px'}),
            
            dcc.Tabs(id="tabs", value='tab-1', children=[
                dcc.Tab(label="📊 Net by Investor Type", value='tab-1', children=[
                    html.Div([
                        dcc.Graph(figure=self.create_net_flow_chart()),
                        self._create_summary_section()
                    ])
                ]),
                
                dcc.Tab(label="📈 Daily Evolution", value='tab-2', children=[
                    html.Div([
                        dcc.Graph(figure=self.create_daily_evolution_chart()),
                        dcc.Graph(figure=self.create_volume_pie_chart())
                    ])
                ]),
                
                dcc.Tab(label="🔍 Sector Analysis", value='tab-3', children=[
                    html.Div([
                        html.H3("Sector Flow Over Time", style={'color': '#003366', 'marginBottom': '20px'}),
                        dcc.Dropdown(
                            id='sector-dropdown',
                            options=[{'label': s, 'value': s} for s in self.evol_sector.columns.get_level_values(1).unique()],
                            value=self.evol_sector.columns.get_level_values(1).unique()[0],
                            style={'marginBottom': '20px'}
                        ),
                        dcc.Graph(id='sector-evolution')
                    ])
                ])
            ])
        ], style={'backgroundColor': '#f8f9fa', 'padding': '20px', 'minHeight': '100vh'})
        
        # Callback para o gráfico de setor
        @app.callback(
            Output('sector-evolution', 'figure'),
            Input('sector-dropdown', 'value')
        )
        def update_sector_chart(selected_sector):
            fig = go.Figure()
            colors = {'LOCALS': '#1f77b4', 'FOREIGNERS': '#ff7f0e', 'RETAIL': '#2ca02c'}
            
            for inv in self.evol_sector.columns.get_level_values(0).unique():
                if (inv, selected_sector) in self.evol_sector.columns:
                    fig.add_trace(go.Scatter(
                        x=self.evol_sector.index,
                        y=self.evol_sector[(inv, selected_sector)],
                        name=inv,
                        mode='lines+markers',
                        line=dict(color=colors.get(inv, '#d62728'), width=2),
                        hovertemplate='<b>%{fullData.name}</b><br>%{x}<br>%{y:.1f}M<extra></extra>'
                    ))
            
            fig.update_layout(
                title=f"Daily Net Flow - {selected_sector}",
                yaxis_title="R$ Million",
                plot_bgcolor='#f8f9fa',
                paper_bgcolor='#ffffff',
                font=dict(family="Arial", size=12),
                title_font=dict(size=16, color='#003366'),
                height=400
            )
            
            return fig
        
        return app
    
    def _create_summary_section(self):
        """Cria seção de resumo"""
        # Calcula top/bottom setores
        top_locals, bottom_locals = self.get_top_bottom_sectors('LOCALS')
        top_foreigners, bottom_foreigners = self.get_top_bottom_sectors('FOREIGNERS')
        
        return html.Div([
            html.H4("📋 Summary", style={'color': '#003366', 'marginBottom': '15px'}),
            
            html.Div([
                html.Div([
                    html.H5("Net Totals (R$ Million)", style={'color': '#003366'}),
                    html.P(f"🏛️ Locals: R$ {self.total_net_by_type['LOCALS']:,.1f}M", 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                    html.P(f"🌍 Foreigners: R$ {self.total_net_by_type['FOREIGNERS']:,.1f}M", 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                    html.P(f"🏪 Retail: R$ {self.total_net_by_type['RETAIL']:,.1f}M", 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),
                
                html.Div([
                    html.H5("Top Sectors (Locals)", style={'color': '#003366'}),
                    html.P("📈 Most Bought: " + ", ".join(top_locals['SECTOR'].head(3)), 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                    html.P("📉 Most Sold: " + ", ".join(bottom_locals['SECTOR'].tail(3)), 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                    html.Br(),
                    html.H5("Top Sectors (Foreigners)", style={'color': '#003366'}),
                    html.P("📈 Most Bought: " + ", ".join(top_foreigners['SECTOR'].head(3)), 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                    html.P("📉 Most Sold: " + ", ".join(bottom_foreigners['SECTOR'].tail(3)), 
                          style={'margin': '5px 0', 'fontSize': '14px'}),
                ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top', 'marginLeft': '4%'})
            ])
        ], style={
            'padding': '20px', 
            'backgroundColor': '#ffffff', 
            'borderRadius': '10px', 
            'marginTop': '20px',
            'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'
        })

# Execução da aplicação
#if __name__ == '__main__':
 #   dashboard = GenialDashboard()
  #  app = dashboard.create_app()
    
    # Para desenvolvimento local
   # print("🚀 Dashboard iniciando...")
    #print("📱 Acesse em: http://localhost:8050")
    #print("🌐 Ou em: http://127.0.0.1:8050")
    #print("⏹️  Para parar: Ctrl+C")
    
    #app.run(debug=True, host='127.0.0.1', port=8050)
    
    # Para produção/compartilhamento (descomente a linha abaixo e comente a de cima)
    # app.run_server(debug=False, host='0.0.0.0', port=8050)

if __name__ == '__main__':
    dashboard = GenialDashboard()
    app = dashboard.create_app()
    import os
    
    # Configuração para produção
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('DEBUG', 'False').lower() == 'true'
    
    if os.environ.get('RENDER'):
        # Produção no Render
        app.run(debug=False, host='0.0.0.0', port=port)
    else:
        # Desenvolvimento local
        print("🚀 Dashboard iniciando...")
        print("📱 Acesse em: http://localhost:8050")
        print("⏹️  Para parar: Ctrl+C")
        app.run(debug=True, host='127.0.0.1', port=8050)