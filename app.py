import streamlit as st
import pandas as pd
import pickle
import os.path
import time

import os.path

from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
import plotly.express as px

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

ConfigSheet = "1ZXU8bC_hPmlpDHLi9t9E0xipk9kX7BG5yoYc1w5Xyts"
ConfigTab = 'Presenters!A1:B20'

st.set_page_config(layout="wide")
reviewers=14

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('secret.json'):
        creds = Credentials.from_authorized_user_file('secret.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('secret.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=ConfigSheet,
                                    range=ConfigTab).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            st.write('%s, %s' % (row[0], row[4]))
    except HttpError as err:
        print(err)


from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

def gsheet_api_check(SCOPES):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

from googleapiclient.discovery import build

def pull_sheet_data(SCOPES, SPREADSHEET_ID, DATA_TO_PULL):
    creds = gsheet_api_check(SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=DATA_TO_PULL).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                  range=DATA_TO_PULL).execute()
        data = rows.get('values')
        print("COMPLETE: Data copied")
        time.sleep(2)
        return data


comDimensions=["Voice", "Speed", "Inspires", "Pauses", "Non-verbal"]
slideDimensions=["Explanation","Interest","Text","Media Quality","Principles"]
prinDimensions=["Coherence","Signaling","Balance","Spatial Contig.","Temporal Contig.","Segmenting","Personalization"]
structDimensions=["Introduction","Background","Design","Support","Summary and Conclusions"]

@st.cache_data
def getDataConfig():
    # Get Configuration
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    data = pull_sheet_data(SCOPES, ConfigSheet, ConfigTab)
    configDf = pd.DataFrame(data[1:], columns=["Name", "Sheet"])
    print("Got configuration")
    return configDf


@st.cache_data
def getDataCommunications(configDf):
    presenters = configDf.shape[0]

    results = []
    observations = []
    for index,row in configDf.iterrows():
        name=row["Name"]
        sheet=row["Sheet"]
        print(name)

        #Getting preference data
        DATA_TO_PULL = 'Instructor' + '!B5:B9'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        selections={}
        ind=0
        for row in data:
            if row[0]=="✅":
                selections[comDimensions[ind]]="Yes"
            elif row[0]=="❌":
                selections[comDimensions[ind]]="No"
            else:
                selections[comDimensions[ind]] = "Unselected"
        #Getting resuts data
        DATA_TO_PULL = 'Instructor' + '!F5:H9'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        values = []
        comments = []
        values.append(name)
        comments.append(name)
        values.append("Instructor")
        values.append("Instructor")
        comments.append("Instructor")
        comments.append("Instructor")
        for row in data:
            if row:
                if row[0] != "":
                    values.append(row[0])
                else:
                    values.append(None)
                if len(row) > 1:
                    comments.append(row[1])
                else:
                    comments.append(None)
        results.append(values)
        observations.append(comments)

        for i in range (1,reviewers+1):
            DATA_TO_PULL = 'Reviewer_'+str(i)+'!F5:H9'
            data = pull_sheet_data(SCOPES,sheet,DATA_TO_PULL)
            if data is not None:
                question=1
                values=[]
                comments=[]
                values.append(name)
                comments.append(name)
                values.append("Reviewer_"+str(i))
                values.append("Reviewers")
                comments.append("Reviewer_" + str(i))
                comments.append("Reviewers")
                print("i")
                print(data)
                for dim in comDimensions:
                    indDim=comDimensions.index(dim)
                    if len(data)>indDim:
                        if data[indDim]:
                            if data[indDim][0] != "":
                                values.append(data[indDim][0])
                            else:
                                values.append(None)
                            if len(data[indDim])>1:
                                comments.append(data[indDim][1])
                            else:
                                comments.append(None)
                        else:
                            values.append(None)
                            comments.append(None)
                    else:
                        values.append(None)
                        comments.append(None)
                results.append(values)
                observations.append(comments)
    print(results)
    print(observations)
    df = pd.DataFrame(results,columns=["Name","Reviewer","Type"]+comDimensions)
    dfComments = pd.DataFrame(observations,columns=["Name","Reviewer","Type"]+comDimensions)
    return selections,df,dfComments

@st.cache_data
def getDataSlides(configDf):
    presenters = configDf.shape[0]

    results = []
    observations = []
    resultsMP =[]
    observationsMP =[]
    for index,row in configDf.iterrows():
        name=row["Name"]
        sheet=row["Sheet"]
        print(name)

        #Getting preference data
        DATA_TO_PULL = 'Instructor' + '!B14:B18'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        selections={}
        ind=0
        for row in data:
            if row[0]=="✅":
                selections[slideDimensions[ind]]="Yes"
            elif row[0]=="❌":
                selections[slideDimensions[ind]]="No"
            else:
                selections[slideDimensions[ind]] = "Unselected"
        #Getting resuts data
        DATA_TO_PULL = 'Instructor' + '!F14:H17'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        values = []
        comments = []
        values.append(name)
        comments.append(name)
        values.append("Instructor")
        values.append("Instructor")
        comments.append("Instructor")
        comments.append("Instructor")
        for row in data:
            if row:
                if row[0] != "":
                    values.append(row[0])
                else:
                    values.append(None)
                if len(row) > 1:
                    comments.append(row[1])
                else:
                    comments.append(None)
        results.append(values)
        observations.append(comments)

        # Getting resuts data
        DATA_TO_PULL = 'Instructor' + '!F19:H25'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        valuesMP = []
        commentsMP = []
        valuesMP.append(name)
        commentsMP.append(name)
        valuesMP.append("Instructor")
        valuesMP.append("Instructor")
        commentsMP.append("Instructor")
        commentsMP.append("Instructor")
        for row in data:
            if row:
                if row[0] != "":
                    valuesMP.append(row[0])
                else:
                    valuesMP.append(None)
                if len(row) > 1:
                    commentsMP.append(row[1])
                else:
                    commentsMP.append(None)
        resultsMP.append(valuesMP)
        observationsMP.append(commentsMP)

        for i in range (1,reviewers+1):
            DATA_TO_PULL = 'Reviewer_'+str(i)+'!F14:H17'
            data = pull_sheet_data(SCOPES,sheet,DATA_TO_PULL)
            if data is not None:
                question=1
                values=[]
                comments=[]
                values.append(name)
                comments.append(name)
                values.append("Reviewer_"+str(i))
                values.append("Reviewers")
                comments.append("Reviewer_" + str(i))
                comments.append("Reviewers")
                print("i")
                print(data)
                for dim in comDimensions:
                    indDim=comDimensions.index(dim)
                    if len(data)>indDim:
                        if data[indDim]:
                            if data[indDim][0] != "":
                                values.append(data[indDim][0])
                            else:
                                values.append(None)
                            if len(data[indDim])>1:
                                comments.append(data[indDim][1])
                            else:
                                comments.append(None)
                        else:
                            values.append(None)
                            comments.append(None)
                    else:
                        values.append(None)
                        comments.append(None)
                results.append(values)
                observations.append(comments)

        for i in range (1,reviewers+1):
            DATA_TO_PULL = 'Reviewer_'+str(i)+'!F19:H25'
            data = pull_sheet_data(SCOPES,sheet,DATA_TO_PULL)
            if data is not None:
                question=1
                valuesMP=[]
                commentsMP=[]
                valuesMP.append(name)
                commentsMP.append(name)
                valuesMP.append("Reviewer_"+str(i))
                valuesMP.append("Reviewers")
                commentsMP.append("Reviewer_" + str(i))
                commentsMP.append("Reviewers")
                print("i")
                print(data)
                for dim in prinDimensions:
                    indDim=prinDimensions.index(dim)
                    if len(data)>indDim:
                        if data[indDim]:
                            if data[indDim][0] != "":
                                valuesMP.append(data[indDim][0])
                            else:
                                valuesMP.append(None)
                            if len(data[indDim])>1:
                                commentsMP.append(data[indDim][1])
                            else:
                                commentsMP.append(None)
                        else:
                            valuesMP.append(None)
                            commentsMP.append(None)
                    else:
                        valuesMP.append(None)
                        commentsMP.append(None)
                resultsMP.append(values)
                observationsMP.append(comments)

    df = pd.DataFrame(results,columns=["Name","Reviewer","Type"]+slideDimensions)
    dfComments = pd.DataFrame(observations,columns=["Name","Reviewer","Type"]+slideDimensions)
    dfMP = pd.DataFrame(resultsMP,columns=["Name","Reviewer","Type"]+prinDimensions)
    dfMPComments = pd.DataFrame(observationsMP,columns=["Name","Reviewer","Type"]+prinDimensions)
    return selections,df,dfComments,dfMP,dfMPComments

@st.cache_data
def getDataStructure(configDf):
    presenters = configDf.shape[0]

    results = []
    observations = []
    for index,row in configDf.iterrows():
        name=row["Name"]
        sheet=row["Sheet"]
        print(name)

        #Getting preference data
        DATA_TO_PULL = 'Instructor' + '!B31:B35'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        selections={}
        ind=0
        for row in data:
            if row[0]=="✅":
                selections[structDimensions[ind]]="Yes"
            elif row[0]=="❌":
                selections[structDimensions[ind]]="No"
            else:
                selections[structDimensions[ind]] = "Unselected"
        #Getting resuts data
        DATA_TO_PULL = 'Instructor' + '!F31:H35'
        data = pull_sheet_data(SCOPES, sheet, DATA_TO_PULL)
        values = []
        comments = []
        values.append(name)
        comments.append(name)
        values.append("Instructor")
        values.append("Instructor")
        comments.append("Instructor")
        comments.append("Instructor")
        for row in data:
            if row:
                if row[0] != "":
                    values.append(row[0])
                else:
                    values.append(None)
                if len(row) > 1:
                    comments.append(row[1])
                else:
                    comments.append(None)
        results.append(values)
        observations.append(comments)

        for i in range (1,reviewers+1):
            DATA_TO_PULL = 'Reviewer_'+str(i)+'!F31:H35'
            data = pull_sheet_data(SCOPES,sheet,DATA_TO_PULL)
            if data is not None:
                question=1
                values=[]
                comments=[]
                values.append(name)
                comments.append(name)
                values.append("Reviewer_"+str(i))
                values.append("Reviewers")
                comments.append("Reviewer_" + str(i))
                comments.append("Reviewers")
                print("i")
                print(data)
                for dim in structDimensions:
                    indDim=structDimensions.index(dim)
                    if len(data)>indDim:
                        if data[indDim]:
                            if data[indDim][0] != "":
                                values.append(data[indDim][0])
                            else:
                                values.append(None)
                            if len(data[indDim])>1:
                                comments.append(data[indDim][1])
                            else:
                                comments.append(None)
                        else:
                            values.append(None)
                            comments.append(None)
                    else:
                        values.append(None)
                        comments.append(None)
                results.append(values)
                observations.append(comments)
    print(results)
    print(observations)
    df = pd.DataFrame(results,columns=["Name","Reviewer","Type"]+structDimensions)
    dfComments = pd.DataFrame(observations,columns=["Name","Reviewer","Type"]+structDimensions)
    return selections,df,dfComments


configDf= getDataConfig()
comSelections, comDf,comDfComments =getDataCommunications(configDf)
slideSelections, slideDf, slideDfComment, MPDf,MPDfComments = getDataSlides(configDf)
structSelections, structDf, structDfComments = getDataStructure(configDf)

st.sidebar.title("Controls")
params=st.experimental_get_query_params()

if "presenter" in params.keys():
    presenterId=int(params["presenter"][0])
    print("Presenter",presenterId)
    presenter=configDf["Name"].values[presenterId]
    print("Presenter Name",presenter)
else:
    presenter=st.sidebar.selectbox(
        'Which student you want to see?',
        options=configDf["Name"].to_list())

dimSelection=st.sidebar.selectbox(
    "Which presentation dimension do you want to see?",
    options=["Communication","Slides","Content & Structure"],
    index=0
)

if dimSelection=="Communication":
    df=comDf
    dfComments=comDfComments
    dimensions=comDimensions
elif dimSelection=="Slides":
    df=slideDf
    dfComments=slideDfComment
    dimensions=["Explanation","Interest","Text","Media Quality"]
elif dimSelection=="Content & Structure":
    df=structDf
    dfComments=structDfComments
    dimensions=structDimensions


dfMelted = pd.melt(df, id_vars=['Name',"Reviewer","Type"], value_vars=dimensions)



mask=dfMelted["Name"]==presenter
filtered=dfMelted[mask]

mask=dfComments["Name"]==presenter
comments=dfComments[mask]



st.title("Presentation Report for "+presenter)
st.header(dimSelection)
dimensionNumber=0
for dim in dimensions:
    if dimensionNumber>0:
        st.markdown("""---""")
    dimensionNumber=dimensionNumber+1
    mask=filtered["variable"]==dim
    justDim=filtered[mask]
    fig = px.strip(justDim,  x="value", width=900, height=400, color="Type",range_x=[1,5.1],stripmode="group")
    fig.update_traces({'marker': {'size': 12}})
    st.subheader(dim)
    st.plotly_chart(fig)
    numberOfComments = 0
    dimIndex = dimensions.index(dim)
    for index, row in comments.iterrows():
        if not ((row[3 + dimIndex] is None) or (row[3 + dimIndex] =="")):
            numberOfComments = numberOfComments + 1
    if numberOfComments>0:
        with st.expander(str(numberOfComments)+" Comments"):
            for index,row in comments.iterrows():
                if not ((row[3 + dimIndex] is None) or (row[3 + dimIndex] =="")):
                    author=row[1]
                    comment=row[3+dimIndex]
                    st.markdown("*"+author+"*: "+comment)
            if numberOfComments==0:
                st.write("No Comments")









#Creating dataset

# for i in range(0,students):
#     responses=[]
#     for j in range (0,len(dimensions)):
#         if results[i][j] is not None:
#
#
#     for i in range(0,len(data)):
#
#     st.write(values)
#     st.write(observations)
    #    data[index]=[i]+row
    #    index=index+1
    #df = pd.DataFrame(data[1:], columns=["Reviewer","Active","Dimension","Low","High","Scores","Notes"])
    #communication=pd.concat([communication,df])

#st.write(communication)
import plotly.subplots as sp
import plotly.graph_objs as go
# filtered['value'] = filtered['value'].replace({"0":"2"})
# filtered['value'] = pd.Categorical(filtered['value'], categories=['1', '2', '3','4','5'])
# dataCounts=filtered.groupby(['value','variable'])["variable"].aggregate('count').unstack()
# #dataCounts=dataCounts.groupby(["value","variable"])["variable"]
# dataDict={}
# dataList=[]
# for index,row in dataCounts.iterrows():
#     valueList=[]
#     for dimension in dimensions:
#         valueList.append(row[dimension])
#     dataDict[index]=valueList
#     dataList.append(valueList)
#
# for values in dataList:
#     valuesSum=0
#     for value in values:
#         valuesSum=valuesSum+value
#
#
#
# data2 = pd.DataFrame(dataDict,index=dimensions)
#
#
# fig, (ax1) = plt.subplots(nrows=1, ncols=1)
# scale=['1',
#  '2',
#  '3',
#  '4',
#  '5']
# #plot_likert.plot_likert(data)
# plot_likert.plot_counts(data2,scale,
#                         plot_percentage=False,  # show absolute values
#                         ax=ax1,  # show on the left-side subplot
#                         legend=1,  # hide the legend for the subplot, we'll show a single figure legend instead,
#                         bar_labels=True, bar_labels_color="snow",
#                         figsize=(5,3)
#                        )
#
#
# # display a single legend for the whole figure
# st.pyplot(fig)

# top_labels = ['Level 1', 'Level 2', 'Level 3', 'Level 4',
#               'Level 5']
#
# colors = ['rgba(38, 24, 74, 0.8)', 'rgba(71, 58, 131, 0.8)',
#           'rgba(122, 120, 168, 0.8)', 'rgba(164, 163, 204, 0.85)',
#           'rgba(190, 192, 213, 1)']
#
# x_data = [[21, 30, 21, 16, 12],
#           [24, 31, 19, 15, 11],
#           [27, 26, 23, 11, 13],
#           [29, 24, 15, 18, 14],
#           [29, 24, 15, 18, 14]]
#
# #x_data = dataList
#
# y_data = dimensions
#
# fig = go.Figure()
#
# for i in range(0, len(x_data[0])):
#     for xd, yd in zip(x_data, y_data):
#         fig.add_trace(go.Bar(
#             x=[xd[i]], y=[yd],
#             orientation='h',
#             marker=dict(
#                 color=colors[i],
#                 line=dict(color='rgb(248, 248, 249)', width=1)
#             )
#         ))
#
# fig.update_layout(
#     xaxis=dict(
#         showgrid=False,
#         showline=False,
#         showticklabels=False,
#         zeroline=False,
#         domain=[0.15, 1]
#     ),
#     yaxis=dict(
#         showgrid=False,
#         showline=False,
#         showticklabels=False,
#         zeroline=False,
#     ),
#     barmode='stack',
#     #paper_bgcolor='rgb(248, 248, 255)',
#     plot_bgcolor='rgb(248, 248, 255)',
#     margin=dict(l=120, r=10, t=140, b=80),
#     showlegend=False,
# )
#
# annotations = []
#
# for yd, xd in zip(y_data, x_data):
#     # labeling the y-axis
#     annotations.append(dict(xref='paper', yref='y',
#                             x=0.14, y=yd,
#                             xanchor='right',
#                             text=str(yd),
#                             font=dict(family='Arial', size=14,
#                                       color='rgb(67, 67, 67)'),
#                             showarrow=False, align='right'))
#     # labeling the first percentage of each bar (x_axis)
#     annotations.append(dict(xref='x', yref='y',
#                             x=xd[0] / 2, y=yd,
#                             text=str(xd[0]) + '%',
#                             font=dict(family='Arial', size=14,
#                                       color='rgb(248, 248, 255)'),
#                             showarrow=False))
#     # labeling the first Likert scale (on the top)
#     if yd == y_data[-1]:
#         annotations.append(dict(xref='x', yref='paper',
#                                 x=xd[0] / 2, y=1.1,
#                                 text=top_labels[0],
#                                 font=dict(family='Arial', size=14,
#                                           color='rgb(67, 67, 67)'),
#                                 showarrow=False))
#     space = xd[0]
#     for i in range(1, len(xd)):
#             # labeling the rest of percentages for each bar (x_axis)
#             annotations.append(dict(xref='x', yref='y',
#                                     x=space + (xd[i]/2), y=yd,
#                                     text=str(xd[i]) + '%',
#                                     font=dict(family='Arial', size=14,
#                                               color='rgb(248, 248, 255)'),
#                                     showarrow=False))
#             # labeling the Likert scale
#             if yd == y_data[-1]:
#                 annotations.append(dict(xref='x', yref='paper',
#                                         x=space + (xd[i]/2), y=1.1,
#                                         text=top_labels[i],
#                                         font=dict(family='Arial', size=14,
#                                                   color='rgb(67, 67, 67)'),
#                                         showarrow=False))
#             space += xd[i]
#
# fig.update_layout(annotations=annotations)
#
# st.plotly_chart(fig)