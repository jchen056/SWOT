import pandas as pd
import numpy as np
import altair as alt
import streamlit as st
from PIL import Image
import streamlit.components.v1 as components
from nvd3 import multiBarChart
from nvd3 import multiBarHorizontalChart
import random

st.header("SWOT Analysis")
with st.expander("Click here to see the orginal SWOT spreadsheet"):
    df_S = pd.read_excel("data/Strength AnalysisData.xlsx")
    df_O = pd.read_excel("data/Opportunity Analysis Data.xlsx")
    df_W = pd.read_excel("data/Weakness Analysis Data.xlsx")
    df_T = pd.read_excel("data/Threat Analysis Data.xlsx")
    df_con = pd.concat([df_S, df_O, df_W, df_T])
    df_con = df_con[['CATEGORY', 'FACTOR TYPE', 'PARAM NAME', 'MIN PROB ADJUSTED VALUE', 'MAX PROB ADJUSTED VALUE', 'AVERAGE PROB ADJUSTED VALUE',
                    'REALISTIC PROB ADJUSTED VALUE', '3 POINT BASED PROB ADJUSTED VALUE', 'PERT BASED PROB ADJUSTED VALUE']]

    df_con.set_axis(['Cat', 'PN', 'Parameter', 'Min', 'Max', 'Avg',
                    'Realistic', '3PT', 'PERT'], axis='columns', inplace=True)
    df_con.reset_index(drop=True)
    df_con.index = pd.RangeIndex(start=0, stop=len(df_con), step=1)
    st.dataframe(df_con)

# Visualization of original data
with st.container():
    html_files = ["htmlFiles/strength.html", "htmlFiles/weakness.html",
                  "htmlFiles/opportunity.html", "htmlFiles/threat.html"]
    tabs = st.tabs(["Strength", "Weakness", "Opportunity", "Threat"])
    for i, x in enumerate(tabs):
        with x:
            p = open(html_files[i])
            components.html(p.read(), height=500, width=800)
            st.caption("multiBarChart from python-nvd3")

# add new data
with st.container():
    st.subheader("SWOT Calculator & Adding New Data")
    tab3, tab4 = st.tabs(["Adding New Data", "SWOT Calculator",])

    with tab4:
        par = st.text_input('Parameter Name', 'i.e. Tech Resources')
        option = st.selectbox(
            'Select a category ',
            ('Strength', 'Weakness', 'Opportunity', 'Threat'))
        if option == "Strength" or option == "Opportunity":
            number = st.number_input(
                'Insert a number for the estimated value in currency', min_value=0, value=400000)
        else:
            number = st.number_input(
                'Insert a number for the estimated value in currency', max_value=0, value=400000)
        values = st.slider(
            'Select a range of Probability %',
            0.0, 100.0, (50.0, 90.0))
        real_prob = st.slider(
            "Select a value for Realistic Prob %", min_value=values[0], max_value=values[1], value=70.0)

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

        df_ans = pd.DataFrame(np.zeros((1, 11)), columns=["Category", "Effect", "Parameter", "3P Prob %", "PERT Prob %",
                                                          "Min", "Max", "Avg", "Realistic", "3P", "PERT"])
        df_ans.loc[0, "Category"] = option
        if option == "Strength" or option == "Opportunity":
            df_ans.loc[0, "Effect"] = "Postive"
        else:
            df_ans.loc[0, "Effect"] = "Negative"
        df_ans.loc[0, "Parameter"] = par
        df_ans.loc[0, ["3P Prob %", "PERT Prob %", "Min", "Max", "Avg", "Realistic",
                       "3P", "PERT"]] = all_calc(number, values[0], real_prob, values[1])
        st.write(df_ans.to_dict('records'))
        df_ans = df_ans[["Category", "Effect", "Parameter", "Min",
                        "Max", "Avg", "Realistic", "3P", "PERT"]]
        df_ans.set_axis(['Cat', 'PN', 'Parameter', 'Min', 'Max', 'Avg', 'Realistic', '3PT',
                        'PERT'], axis='columns', inplace=True)
        # st.write(df_ans)
        # df_con = pd.concat([df_con, df_ans])
        # st.dataframe(df_con)
        # option_Add = st.selectbox(
        #     'Add the new data?',
        #     ('Y', 'N',), index=1)
        # if option_Add == 'Y':
        #     df_con = pd.concat([df_con, df_ans])
        #     df_con.index = pd.RangeIndex(start=0, stop=len(df_con), step=1)
        #     st.dataframe(df_con)
        with tab3:
            if 'data' not in st.session_state:
                data = pd.DataFrame(
                    {'Category': [], 'Parameter': [], 'EstCurrency': [], 'MinProb': [], 'MaxProb': [], 'RealisticProb': []})
                st.session_state.data = data
            data = st.session_state.data
            st.dataframe(data)

            def add_form():
                row = pd.DataFrame({'Category': [st.session_state.input_Category],
                                    'Parameter': [st.session_state.input_Parameter],
                                    'EstCurrency': [st.session_state.input_EstCurrency],
                                    'MinProb': [st.session_state.input_MinProb],
                                    'MaxProb': [st.session_state.input_MaxProb],
                                    'RealisticProb': [st.session_state.input_RealisticProb]
                                    })
                st.session_state.data = pd.concat([st.session_state.data, row])
            dfForm = st.form(key='dfForm')
            with dfForm:
                dfcolumns = st.columns(3)
                with dfcolumns[0]:
                    cat = st.selectbox(
                        'Select a category',
                        ('Strength', 'Weakness', 'Opportunity', 'Threat'),
                        key='input_Category')
                with dfcolumns[1]:
                    st.text_input('Parameter', key='input_Parameter')
                with dfcolumns[2]:
                    st.number_input(
                        'EstCurrency', key="input_EstCurrency")

                dfcolumns1 = st.columns(3)
                with dfcolumns1[0]:
                    st.number_input(
                        'MinProb', min_value=0, max_value=100, key="input_MinProb")
                with dfcolumns1[1]:
                    st.number_input(
                        'MaxProb', max_value=100, key="input_MaxProb")
                with dfcolumns1[2]:
                    st.number_input("RealisticProb",
                                    key="input_RealisticProb")
                st.form_submit_button(on_click=add_form)

            if st.button("Click here to add the new data you have entered"):
                df_newData = data.copy()
                df_newData[["3P Prob %", "PERT Prob %",
                            "Min", "Max", "Avg", "Realistic",
                            "3P", "PERT"]] = 0
                df_newData.index = pd.RangeIndex(
                    start=0, stop=len(df_newData), step=1)
                df_newData.insert(loc=1, column="Effect", value=0)

                for i in range(len(df_newData)):
                    if df_newData.loc[i, "Category"] == "Strength" or df_newData.loc[i, "Category"] == "Opportunity":
                        df_newData.loc[i, "Effect"] = "POSITIVE"
                    else:
                        df_newData.loc[i, "Effect"] = "NEGATIVE"
                    df_newData.loc[i, ["3P Prob %", "PERT Prob %",
                                       "Min", "Max", "Avg", "Realistic",
                                       "3P", "PERT"]] = all_calc(df_newData.loc[i, 'EstCurrency'],
                                                                 df_newData.loc[i,
                                                                                'MinProb'],
                                                                 df_newData.loc[i,
                                                                                'RealisticProb'],
                                                                 df_newData.loc[i, 'MaxProb'])

                df_newData = df_newData[["Category", "Effect", "Parameter", "Min",
                                         "Max", "Avg", "Realistic", "3P", "PERT"]]
                df_newData.set_axis(['Cat', 'PN', 'Parameter', 'Min', 'Max', 'Avg', 'Realistic', '3PT',
                                     'PERT'], axis='columns', inplace=True)
                df_con = pd.concat([df_con, df_newData])
                df_con.index = pd.RangeIndex(start=0, stop=len(df_con), step=1)


df_conMod = pd.DataFrame(np.zeros((len(df_con)*6, 5)), columns=[
                         'Category', 'Effect', 'Parameter', 'Data Points', 'Adjusted Value'])
for i in range(len(df_con)):
    df_conMod.loc[6*i:(6*i+6), 'Parameter'] = df_con.loc[i, 'Parameter']
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

# all the visualization combined
with st.container():
    st.subheader("SWOT Visualization After Adding New Data")
    with st.expander("Click here to see the new SWOT spreadsheet"):
        st.dataframe(df_con)
    swot = ["Strength", "Weakness", "Opportunity", "Threat"]
    dfS = df_conMod[df_conMod["Category"] == swot[0]]
    dfW = df_conMod[df_conMod["Category"] == swot[1]]
    dfO = df_conMod[df_conMod["Category"] == swot[2]]
    dfT = df_conMod[df_conMod["Category"] == swot[3]]

    # xdata = ['Min', 'Max', 'Avg', 'Realistic', '3PT', 'PERT']
    # chartS = multiBarChart(width=800, height=400, x_axis_format=None)
    # chartS.set_containerheader("\n\n<h2>" + "Strength" + "</h2>\n\n")
    # st.dataframe(dfS)

    # output_file = open('Sn.html', 'w')
    # for i in list(set(list(dfS['Parameter']))):
    #     yf = dfS[dfS["Parameter"] == i]['Adjusted Value'].tolist()
    #     yf = [int(i/1000)for i in yf]
    #     chartS.add_serie(
    #         name=i, y=yf, x=xdata)
    # chartS.create_x_axis("Data Points", label="Data Points")
    # chartS.create_y_axis("Adjusted Value", label="Adjusted Value in Thousands")
    # output_file.write(chartS.htmlcontent)
    # output_file.close()
    # p = open('Sn.html')
    # components.html(p.read(), height=500, width=800)

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
                # color=alt.condition(
                #     alt.datum["Adjusted Value"] > 0,
                #     alt.value("steelblue"),  # The positive color
                #     alt.value("orange")  # The negative color
                # ),
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

            # save(chart, 'chart.png')
            # image = Image.open('chart.png', engine="altair_saver")
            # st.image(image, output_format='PNG')
            st.altair_chart(chart, use_container_width=True)
            st.caption("Compact Trellis Grid of Bar Charts from Vega-Altair")
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
