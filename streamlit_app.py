import pandas as pd
import time
import plotly.express as px
import streamlit as st
import matplotlib.colors as mcolors
import openpyxl
from neuralprophet import NeuralProphet
#SET PAGE
st.set_page_config(
    page_title="Sales Dashboard",
    page_icon=":bar_chart:",
    layout="wide",)
#CSS STREAMLIT STYLE
css="""<style>
body{
color:white;}
div.stApp{
background: linear-gradient(0deg, #1F3259 0%, #070B3B 76%);
color: #FEFEFF;
}
/*ScrollBar*/
::-webkit-scrollbar {
    width: 8px; 
    background-color: #f0f0f0; 
}
::-webkit-scrollbar-thumb {
    background-color: #999; 
    border-radius: 4px;
}
::-webkit-scrollbar-track {
    background-color: #ccc; 
}


/*charts*/
.main-svg{
border-radius:10px;
box-shadow: 0 0 2px #070B3B, 0 0 5px #070B3B, 0 0 10px #070B3B;
}
/*sidebar*/
.st-emotion-cache-6qob1r{
background: linear-gradient(0deg, #2C3036 0%, #121820 76%);
}
.st-emotion-cache-2n7b7j,.st-emotion-cache-2n7b7j:hover {
    background: #232956;
}
.st-bo{
background_color:#070B3B}

.st-cr,.st-ck,.st-ck, .st-bk , .st-bb {
background: linear-gradient(0deg, #2C3036 0%, #121820 76%);}

.st-dw, .st-dv, .st-du, .st-dt{
    border-color:#706e6d4f; 
}
.st-e7 {
    border-bottom-color: #212C7B;
}
.st-bo {
    background-color: #212C7B;
}
.st-dn {
    border-bottom-color: #212C7B;
}
.st-dm {
    border-top-color: #212C7B;
}
.st-dl {
    border-right-color:#212C7B;
}
.st-dk {
    border-left-color: #212C7B;
}
.st-b6 {
    color: white;
}
/*map background*/
g.layer.bg{
fill:#121820;
}
span{
color: white;
}

h1, h2{
    color: white;
    -webkit-text-shadow:  0 0 1px #060432, 0 0 5px #060432;
    text-shadow: 0 0 1px #060432, 0 0 5px #060432;
}
p{
    color: white;
    font-size: 25px;
    font-weight:400;
    -webkit-text-shadow:  0 0 1px #070B3B,  0 0 5px #070B3B;
    text-shadow: 0 0 1px #070B3B, 0 0 5px #070B3B;
}
.st-emotion-cache-1xarl3l{
-webkit-text-shadow:  0 0 1px #7a7a7a, 0 0 5px #7a7a7a;
    text-shadow: 0 0 1px #7a7a7a, 0 0 5px #7a7a7a;
}

/*Title*/
.st-emotion-cache-10trblm{
color:white;
box-shadow: 0 0 2px #060432, 0 0 5px #060432, 0 0 20px#060432;
text-align: center;
/*background-color: rgba(18, 24, 32, 0.6); */
}

/*kpi*/
.st-emotion-cache-l9bjmx p {
    font-size: 16px;
    -webkit-text-shadow:  0 0 1px #29283D, 0 0 5px #29283D;
    text-shadow: 0 0 1px #29283D, 0 0 5px #29283D;
    color:#A2A2A2;
    font-weight:500;
}
/*line*/
hr{
    border-bottom: 1px solid #9B9CA6;
    }
/*viewdata*/
.st-emotion-cache-p5msec {
background-color: rgba(44, 48, 54, 0.5); 
box-shadow: 0 0 2px #070B3B, 0 0 5px #070B3B, 0 0 10px #070B3B;
border-radius:20px;
}
.st-emotion-cache-p5msec p{
font-size:1rem;
}
.st-emotion-cache-l9bjmx p{
font-weight:600;
}
svg.eyeqlp51.st-emotion-cache-1pbsqtx.ex0cdmw0{
fill:white;
}

/*download data*/
div.row-widget.stDownloadButton{
    text-align: center;
    align-items: center;
}
button.st-emotion-cache-7ym5gk {
    background-color: #121820;
    border-color: #070B3B;
    color: white;
    box-shadow: 0 0 2px #070B3B, 0 0 5px #070B3B, 0 0 10px #070B3B;
}
.st-emotion-cache-7ym5gk:hover{
    border-color: #070B3B;
    background-color:#212C7B;
    box-shadow: 0 0 2px #2B247F, 0 0 5px #2B247F, 0 0 5px #2B247F;
}
.st-emotion-cache-7ym5gk:focus:not(:active){
    border-color: #070B3B;
}
.st-emotion-cache-7ym5gk p{
    font-size: 1.1rem;
    }
    
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(css,unsafe_allow_html=True)
# READ EXCEL
@st.cache_data
def read_data():
    df = pd.read_excel('Data.xlsx')
    df["Year"] = pd.to_datetime(df['Date'], format='%d.%m.%Y').dt.year
    df["Year_Month"] = pd.to_datetime(df['Date'], format='%d.%m.%Y').dt.to_period('M')
    df["CustomerSatisfaction"] = df["CustomerSatisfaction"].replace({
        "(1) very low": 1,
        "(2) low": 2,
        "(3) ok": 3,
        "(4) high": 4,
        "(5) very high": 5
    })
    return df

df = read_data()
df = df.fillna(0)

# SIDEBAR
with st.sidebar:
    st.title("DASHBOARD SALES")
    st.header("Filters:")

year = st.sidebar.multiselect(
    "Select the Year:",
    options=df["Year"].unique(),
    default=df["Year"].unique()
)

state = st.sidebar.multiselect(
    "Select the State:",
    options=df["State"].unique(),
    default=df["State"].unique()
)

product_type = st.sidebar.multiselect(
    "Select the Product:",
    options=df["Product"].unique(),
    default=df["Product"].unique(),
)

df_selection = df.query(
    "Year == @year & State == @state  & Product == @product_type"
)

# if the dataframe is empty
if df_selection.empty:
    st.warning("No data available based on the current filter settings!")
    st.stop()

# Custom Colors Style
custom_colors= ['#C46400', '#365BDD', '#11208F' ]
custom_colors2= ['#C46400','#3E529B','#11208F']
custom_colors3=['#C46400', '#11208F',]
# TITLE
st.title("📉 Sales Dashboard")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

# METRICS
totalSales = int(df_selection["Revenue"].sum())
avgRating = round(df_selection["CustomerSatisfaction"].mean(), 1)
starRating = "U+2B50" * int(round(avgRating, 0))
avgSaleByTransaction = round(df_selection["Revenue"].mean(), 2)

kpi1, kpi2, kpi3 = st.columns(3)
with kpi1:
    st.metric(label="Total Sales:", value=f"US $ {totalSales:,}")
with kpi2:
    st.metric(label="Average Rating:", value=avgRating, delta="Target: 4")
with kpi3:
    st.metric(label = "Average Sales Per Transaction:", value = f"US $ {avgSaleByTransaction}")
st.markdown("""---""")

#💵 LINECHART Sales By Years(month)
df_selection["Year_Month"] = pd.to_datetime(df_selection["Year_Month"].astype(str), format='%Y-%m')
linechart = pd.DataFrame(df_selection.groupby(df_selection["Year_Month"])["Revenue"].sum()).reset_index()
linechart["Year_Month"] = pd.to_datetime(linechart["Year_Month"])
linechart["Year_Month"] = linechart["Year_Month"].dt.strftime('%Y %b')
lineChart = px.line(linechart,
                    x="Year_Month",
                    y="Revenue",)
lineChart.update_layout(
    plot_bgcolor='#121820',
    paper_bgcolor='#121820',
    xaxis=dict(
        title=dict(text='Month Year', font=dict(color='white')),
        tickmode='linear',
       # tickvals=linechart.index[::3],
       # ticktext=linechart['Year_Month'][::3],
        tickangle=45,
        tickfont=dict(color='white'),
        title_font={"size": 16},
        title_standoff=25),
    yaxis=dict(
        title=dict(text='$Revenue', font=dict(color='white')),
        title_font={"size": 16},),
    title=dict(
        text="Revenue over the years 2017-2019",
        font=dict(color="white", size=22),
        x=0.05,
        y=0.95 ),
    hovermode='closest')
lineChart.update_traces(line=dict(width=3))
linechart['Change'] = linechart['Revenue'].diff()
negative_changes = linechart[linechart['Change'] < 0]
points = px.scatter(x=negative_changes['Year_Month'], y=negative_changes['Revenue'],
                    color=negative_changes['Change'])
lineChart.add_trace(points.data[0])
lineChart.update_traces(marker=dict(color='orange'), selector=dict(type='scatter', mode='markers'), overwrite=True)
st.plotly_chart(lineChart, use_container_width=True)

custom_cmap = mcolors.ListedColormap(['#C46400'] + 3*['#EEEDF4'] + ['#11208F'])
custom_cmap2 = mcolors.ListedColormap(['#C46400'] + 5*['#EEEDF4'] + ['#11208F'])

# Download Data for Sales by Month Year
col1a, col2a = st.columns(2)
with col1a:
    with st.expander("View Data - Sales by Year(Month)"):
        st.write(linechart.style.background_gradient(cmap=custom_cmap2))
        csv = linechart.to_csv(index=False).encode('utf-8')
        st.download_button('Download Data', data = csv, file_name = "DataSalesbyMonthYear.csv",mime = "text/csv",
                           help='Click here to download the data as a CSV file')

revenue_by_state = df_selection.groupby(by=['State'], as_index=False)['Revenue'].sum()
revenue_by_state["State"] = revenue_by_state["State"].replace({
    "Alabama": "AL",
    "Florida": "FL",
    "Georgia": "GA",
    "Mississippi": "MS",
    "North Carolina": "NC",
    "South Carolina": "SC",
    "Tennessee": "TN"
})

# BARCHART - Sales by Product
salesByProduct= df_selection.groupby(by=["Product"], as_index=False)["Revenue"].sum()
col1, col2 = st.columns(2)
with col1:
    #st.subheader("Sales by product")
    barchart = px.bar(salesByProduct,
                      x="Product",
                      y="Revenue",)
    barchart.update_layout(
        plot_bgcolor='#121820',
        paper_bgcolor='#121820',
        legend=dict(
            font=dict(color="white")),
        xaxis=dict(
            title=dict(text="Product", font=dict(color='white')),
            title_font={"size": 16},
            tickfont=dict(color='white')),
        yaxis=dict(
            title=dict(text='$Revenue', font=dict(color='white')),
            title_font={"size": 16},tickfont=dict(color='white')),
        title=dict(
            text=" Sales by Product",
            font=dict(color="white", size=22 ),
            x=0.05, y=0.95),
        hovermode='closest')
    barchart.update_traces(marker=dict(color=['#212C7B', '#212C7B', '#212C7B', '#C46400', '#212C7B']))
    st.plotly_chart(barchart, use_container_width=True)

# MAP CHART Sales by State🌎
with col2:
    mapChart = px.choropleth(revenue_by_state,locations="State",
                             locationmode='USA-states',color="Revenue",
                             color_continuous_scale=custom_colors,
                             scope="usa",
                             title="Total Revenue by State",)
    mapChart.update_layout(coloraxis_colorbar_title="Total Revenue",
                           margin=dict(l=0, r=0, t=0, b=0),
                           paper_bgcolor='#121820',
                           geo=dict(bgcolor=' #121820'),
                           font=dict(
                               color="rgba(255, 255, 255, 1)", size=20),
                           title=dict(
                               text=" Sales by State",
                               font=dict(color="white",size=22),
                               x=0.05, y=0.95),
                           legend=dict(font=dict(color="white")),)
    mapChart.update_geos(
        fitbounds="locations", visible=False,
        showcountries=False,
        showsubunits=False,)
    mapChart.update_xaxes(side="top")
    st.plotly_chart(mapChart, use_container_width=True)

revenue_by_state ["State"] = revenue_by_state ["State"].replace({
    "AL":"Alabama",
    "FL":"Florida",
    "GA":"Georgia",
    "MS": "Mississippi",
    "NC":"North Carolina",
    "SC":"South Carolina",
    "TN":"Tennessee"
})

exp1, exp2 = st.columns(2)
with exp1:
    with st.expander("View Data - Sales By Product"):
        st.write(salesByProduct.style.background_gradient(cmap=custom_cmap))
        csv = salesByProduct.to_csv(index=False).encode('utf-8')
        st.download_button("Download Data", data=csv, file_name="SalesByProduct.csv", mime="text/csv",
                           help='Click here to download the data as a CSV file')
with exp2:
    with st.expander("View Data - Sales By State"):
        st.write(revenue_by_state.style.background_gradient(cmap=custom_cmap2))
        csv = revenue_by_state.to_csv(index=False).encode('utf-8')
        st.download_button("Download Data", data=csv, file_name="SalesByState.csv", mime="text/csv",
                           help='Click here to download the data as a CSV file')

col3, col4 = st.columns(2)
returns_count = df_selection.groupby(by=["Return"], as_index=False)["Revenue"].count()
# PiE CHART - Returns 📩
with col3:
    pieReturns = px.pie(returns_count, values="Revenue", names="Return",
                        hole=0.5, title="Returns")
    pieReturns.update_layout(
        font=dict(
            color="rgba(255, 255, 255, 1)",
            size=20),
        legend=dict(
            font=dict(color="white")),
        title=dict(
            text=" Returns",
            font=dict(color="white", size=22, ),
            x=0.05, y=0.95),
        plot_bgcolor='#121820',
        paper_bgcolor='#121820', )
    pieReturns.update_traces(textfont=dict(color='white'),
                             marker=dict(colors=['#212C7B', '#C46400']))
    st.plotly_chart(pieReturns, use_container_width=True)


delivery_status = df_selection.groupby(by=["Delivery Performance"],as_index=False)["Revenue"].count()
# PIE CHART Delivery 🚚
with col4:
    pieDelivery = px.pie(delivery_status, values='Revenue', names="Delivery Performance",
                         title="Delivery", hole=0.5, )
    pieDelivery.update_layout(
        plot_bgcolor='#121820',
        paper_bgcolor='#121820',
        font=dict(
            color="rgba(255, 255, 255, 1)", size=20),
        legend=dict(font=dict(color="white")),
        title=dict(text=" Delivery on time",
                   font=dict(color="white", size=22),
                   x=0.05, y=0.95), )
    pieDelivery.update_traces(textfont=dict(color='white'),
                              marker=dict(colors=['#C46400', '#212C7B']))
    st.plotly_chart(pieDelivery, use_container_width=True)

filtered_return = df_selection[df_selection['Return'] == 'yes']
returnByProduct = filtered_return.groupby(by=['Product'], as_index=False)['Return'].count()
exp3, exp4 = st.columns(2)
# BARCHART returnByProduct
with exp3:
    with st.expander("More Details - Chart Returns By Product"):
        barchart2 = px.bar(returnByProduct,
                           x="Product",
                           y="Return", )
        barchart2.update_layout(
            plot_bgcolor='#121820',
            paper_bgcolor='#121820',
            legend=dict(
                font=dict(color="white")),
            xaxis=dict(
                title=dict(text="Product", font=dict(color='white')),
                title_font={"size": 16},
                tickfont=dict(color='white')),
            yaxis=dict(
                title=dict(text='Returns', font=dict(color='white')),
                title_font={"size": 16}, tickfont=dict(color='white')),
            title=dict(
                text=" Return by Product",
                font=dict(color="white", size=22),
                x=0.05, y=0.95),
            hovermode='closest')
        barchart2.update_traces(marker=dict(color=['#C46400', '#212C7B', '#C46400', '#212C7B', '#212C7B']))
        st.plotly_chart(barchart2, use_container_width=True)


filtered_delayed = df_selection[df_selection['Delivery Performance'] == 'delayed']
delayByState = filtered_return.groupby(by=['State'], as_index=False)['Delivery Performance'].count()
#BARCHART delayedByState
with exp4:
    with st.expander("More Details - Chart Delivery Delay By State"):
        barchart3 = px.bar(delayByState,
                      x="State",
                      y="Delivery Performance",)
        barchart3.update_layout(
        plot_bgcolor='#121820',
        paper_bgcolor='#121820',
        legend=dict(
            font=dict(color="white")),
        xaxis=dict(
            title=dict(text="Product", font=dict(color='white')),
            title_font={"size": 16},
            tickfont=dict(color='white')),
        yaxis=dict(
            title=dict(text='Returns', font=dict(color='white')),
            title_font={"size": 16}, tickfont=dict(color='white')),
        title=dict(
            text="Delivery delay by State",
            font=dict(color="white", size=22),
            x=0.05, y=0.95),
        hovermode='closest')
        barchart3.update_traces(marker=dict(color=['#212C7B','#212C7B', '#C46400', '#212C7B','#212C7B','#212C7B','#212C7B']))
        st.plotly_chart(barchart3, use_container_width=True)

# CHART with Customer Acquisition Type and Revenue
customer = df_selection.groupby(by=['CustomerAcquisitionType'], as_index=False)['Revenue'].count()
waterfall = px.funnel(customer, x='Revenue',
                      y='CustomerAcquisitionType',
                      title="Effectiveness of customer acquisition type")
waterfall.update_layout(
    plot_bgcolor='#121820',
    paper_bgcolor='#121820',
    title=dict(
        text="Effectiveness of customer acquisition type",
        font=dict(color="white", size=22),
        x=0.05, y=0.95),
    xaxis=dict(
        tickfont=dict(color='white', size=20)),
    yaxis=dict(
        title=dict(text="Customer Acquisition Type", font=dict(color='white')),
        tickfont=dict(color='white')), )
waterfall.update_traces(marker=dict(color=['#212C7B', '#212C7B', '#C46400']))
st.plotly_chart(waterfall, use_container_width=True)


# HEATMAP - State,CustomerAcquisitionType and Revenue
heatmap = px.density_heatmap(df_selection, x="State",
                             y="CustomerAcquisitionType",
                             z="Revenue",
                             title="Sales by State and CustomerAcquisitionType",
                             color_continuous_scale=custom_colors2)
heatmap.update_layout(plot_bgcolor='#121820',
                      paper_bgcolor='#121820',
                      title=dict(
                         text=" Sales by State and Customer Acquisition Type",
                         font=dict(color="white", size=22),
                         x=0.05, y=0.95),
                      xaxis= dict(
                         title=dict(text="State", font=dict(color='white')),
                         tickfont=dict(color='white')),
                      yaxis=dict(
                         title=dict(text="Customer Acquisition Type", font=dict(color='white')),
                         tickfont=dict(color='white')),
                      legend=dict(font=dict(color='white')),)
st.plotly_chart(heatmap, use_container_width=True)

#PREDICTION FOR NEXT YEAR
linechart2 = linechart[['Year_Month','Revenue']]
linechart2.columns = ['ds','y']

progress_text = "Forecast in progress. Please wait..."
my_bar = st.progress(0, text=progress_text)

for percent_complete in range(100):
    time.sleep(0.05)
    my_bar.progress(percent_complete + 1, text=progress_text)
model= NeuralProphet()
model.fit(linechart2)
future = model.make_future_dataframe(linechart2, periods=12)
forecast = model.predict(future)
forecast["ds"] = pd.to_datetime(forecast["ds"])
forecast["ds"] = forecast["ds"].dt.strftime('%Y %b')
actual_prediction= model.predict(linechart2)
actual_prediction["ds"] = pd.to_datetime(actual_prediction["ds"])
actual_prediction["ds"]= actual_prediction["ds"].dt.strftime('%Y %b')
min_date = min(linechart2['ds'])
max_date = max(forecast['ds'])
linechart2.columns=['Date','Actual']
st.success('Forecast ready')
lineChart2 = px.line()
lineChart2.add_scatter( x=linechart2['Date'], y=linechart2['Actual'],mode='lines', name='Actual',line_color='#4328E6')
lineChart2.add_scatter(x=forecast['ds'], y=forecast['yhat1'], mode='lines', name='Forecast', line_color='#FD003D')
lineChart2.add_scatter(x=actual_prediction['ds'], y=actual_prediction['yhat1'], mode='lines', name='Actual Prediction',line_color='#C4C1DB')
lineChart2.update_layout(
    plot_bgcolor='#121820',
    paper_bgcolor='#121820',
    xaxis=dict(
        title=dict(text='Month Year', font=dict(color='white')),
        tickmode='auto',
        tickangle=45,
        tickfont=dict(color='white'),
        title_font={"size": 16},
        title_standoff=25,
        range=[min_date, max_date]),
    yaxis=dict(
        title=dict(text='$Revenue', font=dict(color='white')),
        title_font={"size": 16},
    ),
    title=dict(
        text="Forecast for 2020",
        font=dict(color="white", size=22),
        x=0.05,
        y=0.95),
    legend=dict(font=dict(color="white")),
    hovermode='closest')
lineChart.update_traces(line=dict(width=3))
st.plotly_chart(lineChart2, use_container_width=True)
my_bar.empty()

# Download Dataset
with st.expander("View Data - Whole Data Sales"):
    st.write(df)
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button('Download Data', data = csv, file_name = "Data.csv",mime = "text/csv",
                       help='Click here to download the data as a CSV file')

