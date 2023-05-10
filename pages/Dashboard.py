
import streamlit as st
import pandas as pd
import numpy as np
from nvd3 import multiBarChart
from nvd3 import multiBarHorizontalChart
import altair as alt
import streamlit.components.v1 as components
st.header("SWOT Analysis for New Data")
uploaded_file = st.file_uploader("Choose an Excel file")


def all_calc(est_currency, min_p, r_p, max_p):
    p3_prob = (min_p+r_p+max_p)/3
    pert_prob = (min_p+4*r_p+max_p)/6

    min_value = est_currency*min_p/100
    max_value = est_currency*max_p/100
    avg_value = (min_value+max_value)/2
    real_value = est_currency*r_p/100

    p3_value = (min_value+real_value+max_value)/3
    pert_value = (min_value+4*real_value+max_value)/6
    return p3_prob, pert_prob, min_value, max_value, avg_value, real_value, p3_value, pert_value


if uploaded_file is not None:
    with st.container():
        df = pd.read_excel(uploaded_file)
        tab1, tab2 = st.tabs(["Original Data", "Modified Data"])
        with tab1:
            st.write(df)

        df.dropna(inplace=True)
        for i in range(len(df)):
            est_cur = df.loc[i, "EST. VALUE IN CURRENCY"]
            minP = df.loc[i, "MIN PROB  %"]
            rP = df.loc[i, 'REALISTIC PROB  %']
            maxP = df.loc[i, 'MAX PROB %']
            p3_prob, pert_prob, min_value, max_value, avg_value, real_value, p3_value, pert_value = all_calc(
                est_cur, minP, rP, maxP)
            df.loc[i, "Min"] = min_value
            df.loc[i, "Max"] = max_value
            df.loc[i, "Avg"] = avg_value
            df.loc[i, "Realistic"] = real_value
            df.loc[i, "3PT"] = p3_value
            df.loc[i, "PERT"] = pert_value
        df_con = df[['CATEGORY', 'FACTOR TYPE', 'PARAM NAME',
                    'Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']]
        df_con.set_axis(['Cat', 'PN', 'Parameter', 'Min', 'Max', 'Avg',
                        'Realistic', '3PT', 'PERT'], axis='columns', inplace=True)
        df_con.reset_index(drop=True)
        df_con.index = pd.RangeIndex(start=0, stop=len(df_con), step=1)
        with tab2:
            st.dataframe(df_con)
    with st.container():
        st.subheader("SWOT Visualization")
        swot_cat = ["Strength", "Opportunity", "Weakness", "Threat"]
        html_files = ["htmlFiles/strength1.html",
                      "htmlFiles/opportunity1.html",
                      "htmlFiles/weakness1.html",
                      "htmlFiles/threat1.html"]
        xdata = ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']

        for i in range(len(swot_cat)):
            df_temp = df_con[df_con["Cat"] == swot_cat[i]]
            df_temp.reset_index(drop=True)
            df_temp.index = pd.RangeIndex(start=0, stop=len(df_temp), step=1)
            # st.dataframe(df_temp)
            output_file = open(html_files[i], 'w')
            chart = multiBarChart(width=800, height=400, x_axis_format=None)
            chart.set_containerheader(
                "\n\n<h2>" + swot_cat[i]+": Adjusted Value in K" + "</h2>\n\n")
            for j in range(len(df_temp)):
                tempA = df_temp.loc[j, ['Min', 'Max',
                                        'Avg', 'Realistic', '3PT', 'PERT']]
                tempA = [int(k/1000) for k in list(tempA)]
                chart.add_serie(
                    name=df_temp.loc[j, "Parameter"], y=tempA, x=xdata)
            chart.buildhtml()
            chart.create_x_axis("Data Points", label="Data Points")
            chart.create_y_axis(
                "Adjusted Value", label="Adjusted Value in Thousands")
            output_file.write(chart.htmlcontent)
            output_file.close()
        tS, tW, tO, tT = st.tabs(
            ["Strength", "Weakness", "Opportunity", "Threat"])
        with tS:
            p = open("htmlFiles/strength1.html")
            components.html(p.read(), height=500, width=800)
            st.caption("multiBarChart from python-nvd3")
        with tW:
            p = open("htmlFiles/weakness1.html")
            components.html(p.read(), height=500, width=800)
            st.caption("multiBarChart from python-nvd3")
        with tO:
            p = open("htmlFiles/opportunity1.html")
            components.html(p.read(), height=500, width=800)
            st.caption("multiBarChart from python-nvd3")
        with tT:
            p = open("htmlFiles/threat1.html")
            components.html(p.read(), height=500, width=800)
            st.caption("multiBarChart from python-nvd3")
    with st.container():
        df_conMod = pd.DataFrame(np.zeros((len(df_con)*6, 5)), columns=[
            'Category', 'Effect', 'Parameter', 'Data Points', 'Adjusted Value'])
        for i in range(len(df_con)):
            df_conMod.loc[6*i:(6*i+6),
                          'Parameter'] = df_con.loc[i, 'Parameter']
            df_conMod.loc[6*i:(6*i+6), 'Category'] = df_con.loc[i, 'Cat']
            df_conMod.loc[6*i:(6*i+6), 'Effect'] = df_con.loc[i, 'PN']

            for j in range(6):
                df_conMod.loc[6*i+j, 'Data Points'] = ['Min',
                                                       'Max', 'Avg', 'Realistic', '3PT', 'PERT'][j]

        for i in range(len(df_conMod)):
            param = df_conMod.loc[i, "Parameter"]
            dp = df_conMod.loc[i, 'Data Points']
            df_conMod.loc[i, 'Adjusted Value'] = df_con[df_con["Parameter"] == param][dp].tolist()[
                0]

        swot = ["Strength", "Weakness", "Opportunity", "Threat"]
        dfS = df_conMod[df_conMod["Category"] == swot[0]]
        dfW = df_conMod[df_conMod["Category"] == swot[1]]
        dfO = df_conMod[df_conMod["Category"] == swot[2]]
        dfT = df_conMod[df_conMod["Category"] == swot[3]]
        tab1, tab2 = st.tabs(
            ["Interactive SWOT(horizontal)", "Static SWOT(vertical)"])
        with tab2:
            with st.container():
                df_conModC = df_conMod.copy()
                for i in range(len(df_conMod)):
                    if df_conModC.loc[i, "Category"] == "Strength" or df_conMod.loc[i, "Category"] == "Weakness":
                        df_conModC.loc[i, "Category"] = "Internal"
                    else:
                        df_conModC.loc[i, "Category"] = "External"
                chart = alt.Chart(
                    df_conModC,
                    title="SWOT Analysis",
                    width=60,
                    height=alt.Step(8)).mark_bar().encode(
                    x='Data Points',
                    y='Adjusted Value',
                    color=alt.Color('Parameter'),
                    row="Category",
                    column="Effect",
                    tooltip=['Parameter', "Adjusted Value"],
                ).properties(width=180, height=180).configure_axis(
                    grid=False
                ).configure_title(
                    fontSize=20,
                    anchor='start',
                ).configure_header(
                    titleColor='blue',
                    titleFontSize=15,
                    labelFontSize=14
                ).configure_legend(
                    strokeColor='gray',
                    fillColor='#EEEEEE',
                    cornerRadius=5,
                    orient='top'
                )

                st.altair_chart(chart, use_container_width=True)
                st.caption(
                    "Compact Trellis Grid of Bar Charts from Vega-Altair")
        with tab1:

            with st.container():
                st.caption("Internal Factors")

                chart1 = alt.Chart(dfS, title="Strength").mark_bar().encode(
                    y='Data Points',
                    x='Adjusted Value',
                    color='Parameter',
                    tooltip=['Parameter', "Adjusted Value"],
                    order=alt.Order('Adjusted Value', sort='descending')
                ).interactive()
                # st.altair_chart(chart1, use_container_width=True)

                chart2 = alt.Chart(dfW, title="Weakness").mark_bar().encode(
                    y='Data Points',
                    x='Adjusted Value',
                    color='Parameter',
                    tooltip=['Parameter', "Adjusted Value"],
                    order=alt.Order('Adjusted Value', sort='ascending')
                ).interactive()
                st.altair_chart(chart2 | chart1, use_container_width=True)

            with st.container():
                st.caption("External Factors")

                chart3 = alt.Chart(dfO, title="Opportunity").mark_bar().encode(
                    y='Data Points',
                    x='Adjusted Value',
                    color='Parameter',
                    tooltip=['Parameter', "Adjusted Value"],
                    order=alt.Order('Adjusted Value', sort='descending')
                ).interactive()
                # st.altair_chart(chart3, use_container_width=True)

                chart4 = alt.Chart(dfT, title="Threat").mark_bar().encode(
                    y='Data Points',
                    x='Adjusted Value',
                    color='Parameter',
                    tooltip=['Parameter', "Adjusted Value"],
                    order=alt.Order('Adjusted Value', sort='ascending')
                ).interactive()
                st.altair_chart(chart4 | chart3, use_container_width=True)
                st.caption("Horizontal Stacked Bar Chart from Vega-Altair")
    with st.container():
        st.subheader("Visualization of Summation Data")
        df_Sshort = df_con[df_con["Cat"] == swot[0]]
        df_Wshort = df_con[df_con["Cat"] == swot[1]]
        df_Oshort = df_con[df_con["Cat"] == swot[2]]
        df_Tshort = df_con[df_con["Cat"] == swot[3]]
        for i in ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']:
            df_Sshort.loc["Total", i] = df_Sshort[i].sum()
            df_Wshort.loc["Total", i] = df_Wshort[i].sum()
            df_Oshort.loc["Total", i] = df_Oshort[i].sum()
            df_Tshort.loc["Total", i] = df_Tshort[i].sum()
        df_Sum = pd.DataFrame(np.zeros((4, 7)), columns=[
            'Category', 'Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT'])

        t1, t2 = st.tabs(["Summation data", "Convert to K"])
        with t1:
            for i in range(len(df_Sum)):
                df_Sum.loc[i, 'Category'] = swot[i]
            df_Sum.loc[0, ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']
                       ] = df_Sshort.loc["Total", ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']]
            df_Sum.loc[1, ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']
                       ] = df_Wshort.loc["Total", ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']]
            df_Sum.loc[2, ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']
                       ] = df_Oshort.loc["Total", ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']]
            df_Sum.loc[3, ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']
                       ] = df_Tshort.loc["Total", ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']]
            st.dataframe(df_Sum)
        with t2:
            def divide1000(x):
                return x/1000

            df_Sum[['Min', 'Max', 'Avg',
                    'Realistic', '3PT', 'PERT']] = df_Sum[['Min', 'Max', 'Avg',
                                                           'Realistic', '3PT', 'PERT']].apply(divide1000)
            st.dataframe(df_Sum)

        chartSS = multiBarChart(width=800, height=400, x_axis_format=None)
        chartSS.set_containerheader(
            "\n\n<h2>" + "SWOT Analysis for the Summation Data(Adjusted Value in K)" + "</h2>\n\n")

        output_file = open('htmlFiles/swot.html', 'w')
        for i in range(4):
            chartSS.add_serie(name=swot[i], x=['Min', 'Max', 'Avg',
                                               'Realistic', '3PT', 'PERT'], y=df_Sum.loc[i, ['Min', 'Max', 'Avg',
                                                                                             'Realistic', '3PT', 'PERT']].tolist())
        chartSS.buildhtml()
        chartSS.create_x_axis("Data Points", label="Data Points")
        chartSS.create_y_axis(
            "Adjusted Value", label="Adjusted Value in Thousands")
        output_file.write(chartSS.htmlcontent)
        output_file.close()
        ps = open('htmlFiles/swot.html')
        components.html(ps.read(), height=500, width=800)

        df_PN = pd.DataFrame(np.zeros((2, 7)), columns=["Effects", 'Min', 'Max', 'Avg',
                                                        'Realistic', '3PT', 'PERT'])
        df_PN.loc[0, "Effects"] = "Positive Effect"
        df_PN.loc[0, ['Min', 'Max', 'Avg',
                      'Realistic', '3PT', 'PERT']] = (df_Sum.loc[0, ['Min', 'Max', 'Avg',
                                                                     'Realistic', '3PT', 'PERT']]+df_Sum.loc[2, ['Min', 'Max', 'Avg',
                                                                                                                 'Realistic', '3PT', 'PERT']])
        df_PN.loc[1, "Effects"] = "Negative Effect"
        df_PN.loc[1, ['Min', 'Max', 'Avg',
                      'Realistic', '3PT', 'PERT']] = (df_Sum.loc[1, ['Min', 'Max', 'Avg',
                                                                     'Realistic', '3PT', 'PERT']]+df_Sum.loc[3, ['Min', 'Max', 'Avg',
                                                                                                                 'Realistic', '3PT', 'PERT']])

        df_PN.loc[2, "Effects"] = "Differetial"
        for i in ['Min', 'Max', 'Avg',
                  'Realistic', '3PT', 'PERT']:
            df_PN.loc[2, i] = df_PN[i].sum()

        df_PN_mod = pd.DataFrame(np.zeros((len(df_PN)*6, 3)), columns=[
            'Effect', 'Data Points', 'Adjusted Value'])
        st.dataframe(df_PN)

        for i in range(len(df_PN)):
            df_PN_mod.loc[6*i:(6*i+6), 'Effect'] = df_PN.loc[i, 'Effects']
            for j in range(6):
                df_PN_mod.loc[6*i+j, 'Data Points'] = ['Min',
                                                       'Max', 'Avg', 'Realistic', '3PT', 'PERT'][j]

        for i in range(len(df_PN_mod)):
            param = df_PN_mod.loc[i, "Effect"]
            dp = df_PN_mod.loc[i, 'Data Points']
            df_PN_mod.loc[i, 'Adjusted Value'] = df_PN[df_PN["Effects"] == param][dp].tolist()[
                0]
        # st.dataframe(df_PN_mod)
        chartPN = alt.Chart(df_PN_mod.loc[0:11], title="Positive Effect vs Negative Effect(Adjusted Value in K)").mark_bar().encode(
            x="Adjusted Value",
            y="Data Points",
            color=alt.condition(
                alt.datum["Adjusted Value"] > 0,
                alt.value("black"),  # The positive color
                alt.value("red")  # The negative color
            ),
            tooltip=['Effect', "Adjusted Value"],
        ).interactive()
        chartdf = alt.Chart(df_PN_mod.loc[12:], title="Differential(Adjusted Value in K)").mark_bar().encode(
            x="Adjusted Value",
            y="Data Points",
            color=alt.condition(
                alt.datum["Adjusted Value"] > 0,
                alt.value("black"),  # The positive color
                alt.value("red")  # The negative color
            ),
            tooltip=['Effect', "Adjusted Value"],
        ).interactive()
        st.altair_chart(chartPN & chartdf, use_container_width=True)
