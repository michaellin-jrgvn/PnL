import streamlit as st
import numpy as np
import pandas as pd
import glob
import statsmodels.api as sm
import plotly.express as px
from statsmodels.tsa.seasonal import seasonal_decompose

st.set_page_config(layout="wide")

@st.cache(allow_output_mutation=True)
def read_data(file):
    df = pd.read_excel('./' + file, sheet_name='P&L data',engine='pyxlsb',usecols=['Lookup Code','Description','Accounting Period','Details','Base Amount','T3-Cost Center Analysis Code','Name'])
    df['Base Amount'] = df['Base Amount'] * -1
    return df

@st.cache(allow_output_mutation=True)
def read_2021_data(file):
    df = pd.read_excel('./' + file,usecols=['Lookup Code','Description','Accounting Period','Details','Base Amount','T3-Cost Center Analysis Code','Name'])
    df['Base Amount'] = df['Base Amount'] * -1
    return df

@st.cache
def get_unit_name(df):
    return df.Name.unique()

@st.cache
def get_lookup_code(df):
    return df['Lookup Code'].unique()

@st.cache
def append_all_data():
    df_2020 = read_data('GL 2020 - sent - Dec 2020.xlsb')
    df_2019 = read_data('GL 2019 sent.xlsb')
    df_2021_1 = read_2021_data('GL Jan sent OPS.xlsx')
    df_2021_2 = read_2021_data('GL Feb sent OPS.xlsx')
    df_2021_3 = read_2021_data('GL Mar sent OPS.xlsx')
    df = df_2020.append(df_2019)
    df = df.append(df_2021_1)
    df = df.append(df_2021_2)
    df = df.append(df_2021_3)
    return df

df = append_all_data()

st.title('System Performance')
df_month = df.pivot_table(index=['Lookup Code'], columns=['Accounting Period'], values='Base Amount', aggfunc=np.sum,fill_value=0)
st.dataframe(df_month.style.format('{:,}'))

#df_m_19 = df_2019.pivot_table(index=['Lookup Code'], columns=['Accounting Period'], values='Base Amount', aggfunc=np.sum)
#st.table(df_m_19)

st.title('Unit Headline Performance')
get_unit_name = st.selectbox('select unit name',get_unit_name(df))

@st.cache(allow_output_mutation=True)
def unit_performance(unit_name):
    # Import and pre-process data 
    df_unit_selected = df[df.Name == unit_name]
    df_unit_selected = df_unit_selected.pivot_table(values='Base Amount',columns=['Accounting Period'],index=['Lookup Code'], aggfunc=np.sum,fill_value=0)
    
    # Preprocess data
    df_unit_selected = df_unit_selected.rename({'OTHER FIXED COS':'OTHER FIXED COSTS'})
    
    # Adding MCP, cash UC and UC to the P&L
    df_unit_selected.loc['MCP',:] = df_unit_selected.loc['SALES']+ df_unit_selected.loc['COGS']+df_unit_selected.loc['COST OF LABOR']+df_unit_selected.loc['SEMI']+df_unit_selected.loc['ADVERTISING']
    df_unit_selected.loc['CASH UC',:] = df_unit_selected.loc['MCP']+df_unit_selected.loc['LEASE']+df_unit_selected.loc['ROYALTIES']
    if (df_unit_selected.index == 'OTHER FIXED COSTS').any():
        df_unit_selected.loc['UC',:] = df_unit_selected.loc['CASH UC'] + df_unit_selected.loc['FIXED COSTS'] + df_unit_selected.loc['OTHER FIXED COSTS']
        df_unit_selected = df_unit_selected.reindex(['SALES','COGS','COST OF LABOR','SEMI','ADVERTISING','MCP','ROYALTIES','LEASE','CASH UC','FIXED COSTS','OTHER FIXED COSTS','UC'])
    else:
        df_unit_selected.loc['UC',:] = df_unit_selected.loc['CASH UC'] + df_unit_selected.loc['FIXED COSTS']
        df_unit_selected = df_unit_selected.reindex(['SALES','COGS','COST OF LABOR','SEMI','ADVERTISING','MCP','ROYALTIES','LEASE','CASH UC','FIXED COSTS','UC'])
   
    df_unit_selected = df_unit_selected[df_unit_selected.columns[::-1]]
    
    # Create extra table to show % performance
    df_unit_percent = df_unit_selected.apply(lambda x: x/x.loc['SALES']).mul(100).round(2)
    return df_unit_selected, df_unit_percent

df_unit_abs, df_unit_percent = unit_performance(get_unit_name)
st.subheader('Absolute Value')
st.dataframe(df_unit_abs.style.format('{:,}'))

st.subheader('% over Sales')
st.dataframe(df_unit_percent.style.format('{:,}'))

st.title('Deep Dive')
get_lookup_code = st.selectbox('select lookup code', get_lookup_code(df))

def cost_deepdive(unit_name, lookup_code):
    df_unit_deepdive = df[(df.Name == unit_name) & (df['Lookup Code'] == lookup_code)]
    df_unit_deepdive = df_unit_deepdive.pivot_table(index=['Description'], values='Base Amount', columns = ['Accounting Period'],aggfunc=np.sum,fill_value=0, margins=True)
    df_unit_deepdive = df_unit_deepdive[df_unit_deepdive.columns[::-1]]
    df_unit_deepdive_percent = df_unit_deepdive.apply(lambda x: x/x.loc['All']).mul(100).round(2)
    return df_unit_deepdive, df_unit_deepdive_percent

df_unit_deepdive, df_unit_deepdive_percent = cost_deepdive(get_unit_name, get_lookup_code)
st.dataframe(df_unit_deepdive.style.format('{:,}'))
st.dataframe(df_unit_deepdive_percent.style.format('{:,}'))
st.subheader('Dive Even Deeper')

def get_subitem(unit_name, lookup_code, description, period):
    df_subitem = df[(df.Name == unit_name) & (df['Lookup Code'] == lookup_code) & (df['Description'] == description) & (df['Accounting Period']==period)]
    df_subitem = df_subitem.pivot_table(index=['Accounting Period','Details'], values='Base Amount', aggfunc=np.sum).sort_index(ascending=False)

    return df_subitem

select_details = st.selectbox('Select subitems for further deatails', df_unit_deepdive.index)
select_period = st.selectbox('Select period', df_unit_deepdive.columns)
subitem_df = get_subitem(get_unit_name, get_lookup_code, select_details, select_period)
st.dataframe(subitem_df)

st.title('Item ranking')
select_item_rank = st.selectbox('Select items for ranking', df.Description.unique())
df_ranking = df[df['Description']==select_item_rank]
df_ranking = df_ranking.pivot_table(index=['Name'], values='Base Amount', columns=['Accounting Period'],aggfunc=np.sum,fill_value=0)
df_ranking = df_ranking[df_ranking.columns[::-1]]
df_ranking = df_ranking.sort_values(by=df_ranking.columns[0],ascending=True,axis=0)
st.write(df_ranking)
plt = px.bar(df_ranking, x=df_ranking.index, y=df_ranking.columns[0])
st.plotly_chart(plt, use_container_width=True)


trend = seasonal_decompose(df_unit_abs.loc['SALES',:].T, model='additive')
st.write(trend)