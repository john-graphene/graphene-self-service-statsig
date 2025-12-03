import streamlit as st
import pandas as pd
import os
import math
import numpy as np
import io
from barebones_ver2_ss import main_execute



# @st.cache_data
@st.cache(allow_output_mutation=True)
# @st.cache_data()

def process_file(file):
    if file is not None:
        file_ext=os.path.splitext(file.name)[1]
        # if file_ext==".csv":
        #     df = pd.read_csv(file)
        if file_ext==".xlsx":
            df = pd.read_excel(file,sheet_name=None,na_values=np.nan)

        return df


def input(tab=st,sheet_name="Performance"):
    """ Input """

    global f2
    global f1

    df1 = pd.DataFrame()
    tab.subheader('Upload your files')
    left_upload,right_upload=tab.columns(2)
    if sheet_name == "Performance":
        f1=left_upload.file_uploader(f":file_folder: File ({sheet_name})", type=['xlsx'], accept_multiple_files=False, key=sheet_name, help="Upload the data file sent back by Subbu", on_change=None, args=None, kwargs=None, disabled=False, label_visibility="visible")
    elif sheet_name=="Drivers":
        f1=left_upload.file_uploader(f":file_folder: File ({sheet_name})", type=['xlsx'], accept_multiple_files=False, key=sheet_name, help="Upload the processed driver and equity file", on_change=None, args=None, kwargs=None, disabled=False, label_visibility="visible")
    
    if f1 is not None:
        if sheet_name=="Performance":
        # Process the file and update the dataframe
            df1 = process_file(f1)
            ### Added to handle str vs float
            df1[sheet_name]["Measure Value"]=pd.to_numeric(df1[sheet_name]["Measure Value"], errors='coerce') ## Coerce will change all non-values into np.nan
            df1[sheet_name]["Measure Value"]=df1[sheet_name]["Measure Value"].round(2)                        ## Round to 2 to try prevent floating point issue
            df1[sheet_name] = df1[sheet_name][~df1[sheet_name]['Measure Value'].isnull()]                  ## remove nnp.nan value in Measure value column, to try to tackle pivot issue
            ### END Added to handle str vs float

        # base=df1["Base"]
            performance=df1[sheet_name]
            df1=performance.copy()

        ## Remove Base
            # df1=performance.merge(base,how='left',on=['PeriodKey', 'Category',
        #'Subcategory', 'Country', 'Segment', 'Brand'],suffixes=("","_right"))
            df1=df1[['Category','Subcategory','Segment', 'Country', 'Brand',
            'Type', 'Subtype', 'Content', 'Measure Value' ]]
        elif sheet_name=="Drivers":
            df1 = process_file(f1)
            df1 = df1[sheet_name]

        # if tab!=None:
        tab.write(df1)
        # else:
        #     df1
        

    return df1,f1

def get_unique(df1):
    global cat_list,subcat_list,country_list,type_list,subtype_list,brand_list_segment_list
    # unique_list=list(df[unique_col].unique())
    cat_list=list(df1["Category"].unique())
    subcat_list=list(df1["Subcategory"].unique())
    country_list=list(df1["Country"].unique())



    type_list=list(df1["Type"].unique())

    try:
        subtype_list=list(df1[df1["Type"]==selected_type]["Subtype"].unique())
    except:
        subtype_list=["None"]
    brand_list=["None"]
    segment_list=["None"]


def get_benchmark(df_local,split,types):
    if types=="Max":
        return ["None"]
    elif split=="Brand":
        return list(df_local[(df_local["Type"]==selected_type) & (df_local["Subtype"].isin(selected_subtype))]["Brand"].unique())
    elif split=="Segment":
        return df_local[(df_local["Type"]==selected_type) & (df_local["Subtype"].isin(selected_subtype))]["Segment"].unique()

 
def select_box(df_local,tab=st):
    global selected_cat,selected_subcat,selected_country
    global selected_type,selected_subtype,selected_split,unique_split,selected_statsig_type,benchmark_target
    tab.subheader("Filter data")
    sel_cat,sel_subcat,sel_country=tab.columns(3)

    sel_cat_help="Select Category to be filtered to"
    sel_subcat_help="Select Subcategory to be filtered to"
    sel_cty_help="Select Country to be filtered to"
    sel_type_help="Select Type to be filtered to"
    sel_subtype_help="Select Subtype to be filtered to"
    sel_split_help="Select comparison type, cross segment or cross brand comparison"
    sel_statsig_help="Select type of statsig, Benchmark or Max logic"
    sel_benchmark_help="Select Benchmark Brand/Segment (Appicable to Statsig method = benchmark only)"


    selected_cat=sel_cat.selectbox("Category",cat_list,help=sel_cat_help)
    selected_subcat=sel_subcat.selectbox("Subcategory",df_local[df_local["Category"]==selected_cat]["Subcategory"].unique(),help=sel_subcat_help)
    selected_country=sel_country.selectbox("Country",df_local[(df_local["Category"]==selected_cat) & (df_local["Subcategory"]==selected_subcat)]["Country"].unique(),help=sel_cty_help)

    df_local=df_local[(df_local["Category"]==selected_cat) & (df_local["Subcategory"]==selected_subcat) &  (df_local["Country"]==selected_country)]


    sel_type,sel_subtype,sel_split=tab.columns(3)

    selected_type=sel_type.selectbox("Type",type_list,help=sel_type_help)
    # selected_subtype=sel_subtype.selectbox("Subtype",df_local[df_local["Type"]==selected_type]["Subtype"].unique(),help=sel_subtype_help)
    selected_subtype=sel_subtype.multiselect("Subtype",df_local[df_local["Type"]==selected_type]["Subtype"].unique(),help=sel_subtype_help)

    # selected_split=sel_split.selectbox("Split",split_list,help=sel_split_help)
    selected_split="Segment"

    statsig_type,bencmark_target,empty=tab.columns(3)
    selected_statsig_type=statsig_type.selectbox("Statsig Method",["Benchmark","Max"],help=sel_statsig_help)
    benchmark_target=bencmark_target.selectbox("Benchmark",get_benchmark(df_local,selected_split,selected_statsig_type),help=sel_benchmark_help)
    
    ## Find the unique list of split
    unique_split=get_benchmark(df_local,selected_split,"Benchmark")
    return df_local


def find_largest_and_second_largest(numbers_list):
    largest = None
    second_largest = None

    numbers_list = [num for num in numbers_list if type(num)!=str] ## Remove all string from the number list - e.g "-"

    for num in numbers_list:
        ## Added to ignore "-"

        ## END Added to ignore "-"

        # st.write(type(num))  

        if not math.isnan(num):
            num=int(round(num,0))

            # st.write("hello " + str(num))
            if largest is None or num > largest:
                second_largest = largest
                largest = num
            elif ((second_largest is None or num > second_largest) and  int(num)!=int(largest)):
                second_largest = num

    return largest, second_largest

def find_threshold(number,base=None):

    if ((type(number) == str) or (number is None)):
        return "Error"
    elif (base==None) or (base>1100):
        if number>45:
            return (4.1*1.25)
        elif number >30:
            return (1.76*1.25)
        elif number >15:
            return (0.98*1.25)
        elif number > 0:
            return (0.56*1.25)
        else:
            return "Error"
    elif (base>750):
        if number>45:
            return (4.7*1.25)
        elif number >30:
            return (2.21*1.25)
        elif number >15:
            return (1.73*1.25)
        elif number > 0:
            return (0.97*1.25)
        else:
            return "Error"
    elif (base>300):
        if number>45:
            return (6.3*1.25)
        elif number >30:
            return (3.22*1.25)
        elif number >15:
            return (2.2*1.25)
        elif number > 0:
            return (1.23*1.25)
        else:
            return "Error"
    else:
        return "Base less than 300!!"




    
def output(df,name=None,tab=None):
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False,na_rep="-")           
    
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()
        if name==None:
            try:
                name=selected_cat[:10]+"_"+selected_subcat[:10]+"_"+selected_country[:3]+"_"+selected_type+" - "+ selected_subtype[0] + "_"+ selected_split + "_"+ selected_statsig_type+"_"+str(base) + ".xlsx"
            except:
                name="default.xlsx"
        if tab==None:
            st.download_button(
                label="Download Excel worksheets",
                data=buffer,
                file_name=name,
                mime="application/vnd.ms-excel"
            )
        else:
            tab.download_button(
                label="Download Excel worksheets",
                data=buffer,
                file_name=name,
                mime="application/vnd.ms-excel"
            )

           
def output_pptx(df,statsig="max",base=1200,name=None,tab=None):
    df=df.dropna(how='all')
    df = df.dropna(axis=1, how='all')

    statsig=statsig.lower()
    buffer = main_execute(df,statsig,base,"Segment/Brand 1")
    if name==None:
        try:
            name=selected_cat[:10]+"_"+selected_subcat[:10]+"_"+selected_country[:3]+"_"+selected_type+" - "+ selected_subtype[0] + "_"+ selected_split + "_"+ selected_statsig_type+"_"+str(base) + ".pptx"
        except:
            name="default.pptx"

    if tab==None:
        tab = st



    tab.download_button(
        label="Download PPT",
        data=buffer,
        file_name=name

    )


def multi_select_and_df(df1,tab=st):
    # global base,dataframe_show,df_pivot
    split_list, dataframe_show = tab.columns([1, 3])
    
    ### Multi select
    i=1
    selected_segment=[]

    split_list.header("Segment List")
    # split_list.header("Segment/Brand List")

    base_tooltip="Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""
    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}
    def format_func(option):
        return base_dict[option]

    base = split_list.selectbox("Base", options=list(base_dict.keys()), format_func=format_func,help=base_tooltip)
    # base=split_list.number_input("Base",301,10000,1200,50,help=base_tooltip)
    if selected_statsig_type=="Benchmark":

        for seg_brand in unique_split:
            if seg_brand!=benchmark_target:
                seg_chekced=split_list.checkbox(seg_brand,key="seg_brand"+str(i))
                i=i+1
                if seg_chekced:
                    selected_segment.append(seg_brand)
    else:
        for seg_brand in unique_split:
            seg_chekced=split_list.checkbox(seg_brand,key="seg_brand"+str(i))
            i=i+1
            if seg_chekced:
                selected_segment.append(seg_brand)


    ### Multi select
    dataframe_show.header("Statistical Significance table")
    ## Index
    index=list(df1.columns)
    index.remove("Measure Value")


    index.remove(selected_split)

    # index.remove("No_of_People")

    index.remove("Subtype")


    ## Filter df
    if selected_statsig_type!="Benchmark":
        if selected_split=="Brand":
            filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype))& (df1["Brand"].isin(selected_segment))]
        elif selected_split=="Segment":
                filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype))& (df1["Segment"].isin(selected_segment))]
    else:        
        if selected_split=="Brand":
            filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype))& (df1["Brand"].isin(selected_segment+[benchmark_target]))]
        elif selected_split=="Segment":
                filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype))& (df1["Segment"].isin(selected_segment+[benchmark_target]))]

    ## Pivot

    try:
        # df_pivot=filtered_df.pivot(index=index,columns=selected_split,values="Measure Value")
        # df_pivot=df_pivot.reset_index()
        index.remove("Brand")
        df_pivot=filtered_df.pivot(index=index,columns=selected_split,values="Measure Value")
        df_pivot=df_pivot.reset_index()

        if selected_statsig_type !="Benchmark":
            if len(selected_subtype) < 1:
                split_list.write("Please select 1 or more subtype")
            else:
             
                try:
                    df_pivot=df_pivot[index+selected_segment] ## Sort columns
                except:
                            split_list.write("<span style='font-size:20px;padding-left: 5px;'> :exclamation: :exclamation:  Pivot Error :exclamation:  :exclamation: </span>", unsafe_allow_html=True)

                       
        elif selected_statsig_type=="Benchmark": 
            try:
                df_pivot=df_pivot[index+[benchmark_target]+selected_segment] ## Sort columns
            except:
                
                if len(selected_subtype) >= 1:
                    split_list.write("<span style='font-size:20px;padding-left: 5px;'> :exclamation: :exclamation:  Pivot Error :exclamation:  :exclamation: </span>", unsafe_allow_html=True)
                else:
                     split_list.write(":white_frowning_face: Please select 1 or more subtype")
        # df_pivot=df_pivot.drop(columns=["Country","Category","Subcategory","Type","Subtype"])
        df_pivot=df_pivot.drop(columns=["Country","Category","Subcategory","Type"])


      
        df_pivot_styler=apply_statsig(df_pivot,selected_statsig_type,base)
        df_pivot_styler_formatted = df_pivot_styler.format(precision=1, na_rep='-')
 

        ## To fill na with "-"
        # df_pivot = df_pivot.fillna("")

        # df_pivot = df_pivot.fillna(None)
        ## END To fill na with "-"
        dataframe_show.dataframe(df_pivot_styler_formatted)


        # dataframe_show.write("Error")


          ## To fill "" with "-" for output
        # df_pivot = df_pivot.applymap(lambda x: st.write(x))

        output(df_pivot,tab=dataframe_show)
    except Exception as Esc:
        if ((not df_pivot.empty) and (len(selected_segment) >=1)):

            dataframe_show.write("<span style='font-size:20px;padding-left: 10px;'> :exclamation: :exclamation:  DATA ERROR :exclamation:  :exclamation: </span>", unsafe_allow_html=True)
            dataframe_show.write(Esc)
        else:
            dataframe_show.write(":white_frowning_face: Please select at least 1 segment from the Segment list")


def apply_statsig(df_pivot,selected_statsig_type,base,first_col=2):
    def max_logic(row,format,first_col):
        values=row[first_col:]
        highlight=None
        largest,sec_largest=find_largest_and_second_largest(values)

        threshold=find_threshold(largest,base)
        # st.write(row,largest,sec_largest,threshold)
        if largest is not None and sec_largest is not None:
            if largest-sec_largest>threshold:
                highlight=largest
        # elif largest is not None:
        #     highlight=largest

        return_list=[]
        for col in row:
            try:
                if int(round(col,0))==int(round(highlight,0)):
                    return_list.append(format)
                else:
                    return_list.append('')
            except:
                return_list.append('')


        return return_list
    def benchmark_logic(row,sup_format,inf_format,first_col):
        benchmark=row[first_col]
        values=row[first_col+1:]
        return_list=[0]*(first_col+1) ## Indexes and benchmark column 
        format_return_list=[]
        for value in values:
            threshold=find_threshold(value,base)
            if threshold=="Error":
                return_list.append(0)
            elif ((type(value) == str) or (type(benchmark) == str)):   ## to handle when either value or benchmark is "-"
                return_list.append(0)
            elif value-benchmark>threshold:
                return_list.append(1)
            elif value-benchmark<-threshold:
                return_list.append(-1)
            else:
                return_list.append(0)

        for value in return_list:
            if value ==0:
                format_return_list.append("")
            elif value==1:
                format_return_list.append(sup_format)
            elif value==-1:
                format_return_list.append(inf_format)
  
        return format_return_list
    
  

    if selected_statsig_type=="Max":
        df_pivot=df_pivot.style.apply(lambda x:max_logic(x,'color: green;background-color:lightgreen',first_col),axis=1)
    elif selected_statsig_type=="Benchmark":
        df_pivot=df_pivot.style.apply(lambda x:benchmark_logic(x,'color: red;background-color:pink','color: green;background-color:lightgreen',first_col),axis=1)


        
    return df_pivot



def empty_df(tab=st):
    num_rows = 50
    num_columns = 20

    columns =["Content"] + ["Segment/Brand {}".format(i) for i in range(1, num_columns+1)]
    df = pd.DataFrame(index=range(num_rows), columns=columns)
    updated_df=tab.data_editor(df)
    return updated_df

def statsig_tab_sel_box(tab=st):


    sel_statsig_help="Select type of statsig, Benchmark or Max logic"
    statsig_Type=tab.selectbox("Statsig Method",["Benchmark","Max"],help=sel_statsig_help,key=str(tab)+"_sstype")

    base_tooltip="Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""
    
    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}
    def format_func(option):
        return base_dict[option]

    base = tab.selectbox("Base (Overall Segment with Brand = All)", options=list(base_dict.keys()), format_func=format_func,help=base_tooltip)
    
    # base=tab.number_input("Base (Overall Segment with Brand = All)",301,10000,1200,50,help=base_tooltip,key=str(tab)+"_base")

    return statsig_Type,base

def statsig_tab_highlight(updated_df,ss_type,base,tab=st):

    #   Convert selected columns to float
    # columns_to_convert = updated_df.columns[1:]

    # updated_df[columns_to_convert] = updated_df[columns_to_convert].astype(float)

    for column in updated_df.columns[1:]:

        updated_df[column] = pd.to_numeric(updated_df[column].str.replace('%', ''), errors='coerce')

    updated_df=apply_statsig(updated_df,ss_type,base,first_col=1)

#   Chane from 1 dp to no dp
    updated_df = updated_df.format(precision=1, na_rep='-')
    # updated_df = updated_df.format(precision=0, na_rep='-')


    return updated_df


def dande_statsig_select(df,segment,tab=st):

    sel_statsig_help="Select type of statsig, Benchmark or Max logic"
    sel_benchmark_help="Select Benchmark Brand/Segment (Appicable to Statsig method = benchmark only)"
    statsig_type,bencmark_target,base_sel=tab.columns(3)
    selected_statsig_type=statsig_type.selectbox("Statsig Method",["Benchmark","Max"],help=sel_statsig_help,key=segment+"_type")

    if selected_statsig_type=="Max":
        df_selection=[None]
    elif selected_statsig_type=="Benchmark":
        df_selection=df[df["Segment"]==segment]["Brand"].unique()
    benchmark_target=bencmark_target.selectbox("Benchmark",df_selection,help=sel_benchmark_help,key=segment+"_target")

    
    base_tooltip="Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""
    
    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}
    def format_func(option):
        return base_dict[option]

    base = base_sel.selectbox("Base (Overall Segment with Brand = All)", options=list(base_dict.keys()), format_func=format_func,help=base_tooltip,key=segment)
    

    return selected_statsig_type,benchmark_target,base
def dande_segment(df,tab=st):
    for segment in df["Segment"].unique():
        tab.subheader(segment)

        ## Select box for statsig method and benchmarl
        dande_seleceted_statsig,dande_benchmark,base=dande_statsig_select(df,segment,tab)

        ## Get CBI
        df_cbi=df.copy()
        df_cbi=df_cbi[df_cbi["Segment"]==segment]
        df_cbi=df_cbi.drop(columns=["Driver_Score","Driver","Equity_Score"])
        df_cbi=df_cbi.drop_duplicates()
        df_cbi["Driver"] = "CBI"
        df_cbi=df_cbi.rename(columns={"CBI":"Equity_Score"})
        df_with_cbi = pd.concat([df[df["Segment"]==segment],df_cbi])


        ## Pivot
        df_pivot=df_with_cbi[df_with_cbi["Segment"]==segment].pivot(index=["Driver_Score","Driver"],columns="Brand",values="Equity_Score").reset_index()
        df_pivot=df_pivot.sort_values("Driver_Score",ascending = False)

        ## Rearranfe column if is benchamrk
        if dande_seleceted_statsig == "Benchmark":
            other_brand=df[df["Segment"]==segment]["Brand"].unique().tolist()
            other_brand.remove(dande_benchmark)
            # other_brand=other_brand.to_list()
            # tab.dataframe(other_brand)

            df_pivot=df_pivot[["Driver_Score","Driver",dande_benchmark]+other_brand]


        updated_df=apply_statsig(df_pivot,dande_seleceted_statsig,base,first_col=2)

        tab.dataframe(updated_df)
        output(updated_df,name="D&E_"+segment+"xlsx",tab=tab)

    pass

def app():
    st.title('Self service Statistical Significance app')
    # allow_cors()

    global buffer
    buffer = io.BytesIO()
    statsig_tab,pop_tab=st.tabs(["Statistical Significant",'PoP Statistical Significant'])

    ## Statsig _tab
    statsig_tab.write("If **Benchmark** is the selected statsig method, **first column** would be assumed to be the benchmark value")
    updated_df=empty_df(statsig_tab)
    ss_type,base=statsig_tab_sel_box(statsig_tab)
    applied_df=statsig_tab_highlight(updated_df,ss_type,base,statsig_tab)
    statsig_tab.subheader("Results")
    statsig_tab.dataframe(applied_df)
    pop_tab.subheader("Work in progress")
    output_pptx(updated_df,statsig=ss_type,base=base,tab=statsig_tab)

    # ## D&E _tab
    # df2,f2=input(dande_tab,sheet_name="Drivers")
    # if f2 is not None:
    #     dande_segment(df2,dande_tab)


    # # Performance tab
    # df1,f1=input(fixed_tab)


    # if f1 is not None:
    #     global split_list
    #     split_list=["Segment","Brand"]
    #     get_unique(df1)
    #     df1=select_box(df1,fixed_tab)
    #     multi_select_and_df(df1,fixed_tab)

    
    # if (f1 is None):
    #     return

if __name__ == '__main__':
    st.set_page_config(page_title="Self Service statsig", page_icon=":chart_with_upwards_trend:", layout="wide")
    
    app()
