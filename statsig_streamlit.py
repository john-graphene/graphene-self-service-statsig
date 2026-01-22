import streamlit as st
import pandas as pd
import os
import math
import numpy as np
import io
from barebones_ver2_ss import main_execute
import warnings
warnings.filterwarnings('ignore')

# @st.cache_data
# @st.cache(allow_output_mutation=True)
# @st.cache_data()
# @st.cache_resource()

###### Building Functions ######
def empty_df(tab=st):
    """Statistical Significant Tab: Implement empty DataFrame for data input."""

    num_rows = 50
    num_columns = 20

    columns =["Content"] + ["Segment/Brand {}".format(i) for i in range(1, num_columns+1)]
    df = pd.DataFrame(index=range(1,num_rows+1), columns=columns)
    updated_df=tab.data_editor(df)
    return updated_df


def statsig_tab_sel_box(tab=st):
    """Statistical Significant Tab: Implement drop-down boxes for selecting Stat-sig type and Base range."""
    sel_statsig_help="Select type of statsig, Benchmark or Max logic. Benchmark (Inverse) is same as Benchmark but to be used for when inverse stat-sig logic is needed, e.g. Negative Themes, Dissatisfiers, Barriers, etc."
    statsig_Type=tab.selectbox("Statsig Method",["Benchmark","Max","Benchmark (Inverse)"],help=sel_statsig_help,key=str(tab)+"_sstype")

    base_tooltip="Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""

    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}
    def format_func(option):
        return base_dict[option]

    base = tab.selectbox("Base (Overall Segment with Brand = All)", options=list(base_dict.keys()), format_func=format_func,help=base_tooltip)

    # base=tab.number_input("Base (Overall Segment with Brand = All)",301,10000,1200,50,help=base_tooltip,key=str(tab)+"_base")

    return statsig_Type,base


def empty_pop_df(tab=st):
    """PoP Statistical Significant Tab: Implement empty DataFrame for data input."""
    num_rows = 30
    num_columns = 8

    columns =["Content"] + ["Period {}".format(i) for i in range(1, num_columns+1)]
    df = pd.DataFrame(index=range(1,num_rows+1), columns=columns)
    updated_pop_df = tab.data_editor(df)
    return updated_pop_df


def pop_statsig_tab_sel_box(tab=st):
    """PoP Statistical Significant Tab: Implement drop-down boxes for selecting Base range."""
    base_tooltip="Pls input No_of_People from Base sheet for the previous period. Pls adjust according to desired period-on-period of comparison, i.e. Period 3 vs Period 2."

    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}
    def format_func(option):
        return base_dict[option]
    pop_base = tab.selectbox("Previous Period Base (Overall Segment with Brand = All)", options=list(base_dict.keys()), format_func=format_func,help=base_tooltip, key="PoP_base")

    # base=tab.number_input("Base (Overall Segment with Brand = All)",301,10000,1200,50,help=base_tooltip,key=str(tab)+"_base")

    return pop_base


def process_file(file):
    """Read file input, returns DataFrame."""
    if file is not None:
        file_ext=os.path.splitext(file.name)[1]
        if file_ext==".csv":
            df = pd.read_csv(file)
        if file_ext==".xlsx":
            df = pd.read_excel(file,sheet_name=None,na_values=np.nan)

        return df


def get_file_input(tab=st,sheet_name="Performance"):
    """ Implement Button for File Input, Clean-up DataFrame. """
    # global f2
    # global f1
    df1 = pd.DataFrame()
    tab.subheader('Upload your files')
    left_upload,right_upload=tab.columns(2)
    if sheet_name == "Performance":
        f1=left_upload.file_uploader(f":file_folder: File ({sheet_name})", type=['xlsx'], accept_multiple_files=False, key=sheet_name, help="Upload the data file sent back by Subbu", on_change=None, args=None, kwargs=None, disabled=False, label_visibility="visible")
    elif sheet_name=="Drivers":
        f1=left_upload.file_uploader(f":file_folder: File ({sheet_name})", type=['CSV','xlsx'], accept_multiple_files=False, key=sheet_name, help="Upload the processed driver and equity file", on_change=None, args=None, kwargs=None, disabled=False, label_visibility="visible")
    
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
            df1=df1[['PeriodKey','Category','Subcategory','Segment', 'Country', 'Brand',
            'Type', 'Subtype', 'Content', 'Measure Value' ]]
        elif "driver" in sheet_name.lower():
            df1 = process_file(f1)
            if type(df1) == dict:
                df1 = df1[sheet_name]
            df1 = df1.drop(columns=['PeriodDateEnd','Month','Quarter','Year'], errors='ignore')

        # if tab!=None:
        tab.dataframe(df1, hide_index=True)
        # else:
        #     df1
        

    return df1,f1


def get_unique(df1):
    """Performance Tab: Get unique values for PeriodKey, Category, Subcategory, Country, Type, Subtype columns"""
    global periodkey_list,cat_list,subcat_list,country_list,type_list,subtype_list#,brand_list_segment_list
    # unique_list=list(df[unique_col].unique())
    periodkey_list=list(df1["PeriodKey"].unique())
    cat_list=list(df1["Category"].unique())
    subcat_list=list(df1["Subcategory"].unique())
    country_list=list(df1["Country"].unique())
    type_list=list(df1["Type"].unique())
    try:
        # subtype_list=list(df1[df1["Type"]==selected_type]["Subtype"].unique())
        subtype_list=list(df1[df1["Type"].isin(selected_type)]["Subtype"].unique())
    except:
        subtype_list=["None"]
    # brand_list=["None"]
    # segment_list=["None"]


def get_benchmark(df_local,split,statsig_type):
    """Performance Tab: Get list of unique Brands if Segment Comparison, or unique Segments of Brand Comparison."""
    if statsig_type=="Max":
        return ["None"]
    elif split=="Brand":
        return list(df_local[(df_local["Type"].isin(selected_type)) & (df_local["Subtype"].isin(selected_subtype))]["Brand"].unique())
        # return list(df_local[(df_local["Type"]==selected_type) & (df_local["Subtype"].isin(selected_subtype))]["Brand"].unique())
    elif split=="Segment":
        return df_local[(df_local["Type"].isin(selected_type)) & (df_local["Subtype"].isin(selected_subtype))]["Segment"].unique()
        # return df_local[(df_local["Type"]==selected_type) & (df_local["Subtype"].isin(selected_subtype))]["Segment"].unique()

 
def performance_select_box(df_local,tab=st):
    """Performance Tab: Implement drop-down boxes for selecting multiple filters."""
    global selected_cat,selected_subcat,selected_country,selected_segment,selected_brand
    global selected_type,selected_subtype,selected_split,unique_split,selected_statsig_type,benchmark_target
    tab.subheader("Select options below.")

    sel_periodkey_help = "Select PeriodKey to be filtered to"
    sel_cat_help="Select Category to be filtered to"
    sel_subcat_help="Select Subcategory to be filtered to"
    sel_cty_help="Select Country to be filtered to"
    sel_type_help="Select Type to be filtered to"
    sel_split_help="Select comparison type, cross segment or cross brand comparison"
    sel_subtype_help="Select Subtype to be filtered to"
    sel_segment_help="Select single Segment from list which has multiple Brands for comparison."
    sel_brand_help="Select single Brand from list which has multiple Segments for comparison."
    sel_statsig_help="Select type of statsig, Benchmark or Max logic."
    sel_benchmark_help="Select Benchmark Brand/Segment (Leave as blank for Max Statsig Method)"

    tab.markdown("**Step 1:** Choose PeriodKey, Category, Subcategory, Country fixed filters.")
    sel_periodkey,sel_cat,sel_subcat,sel_country=tab.columns(4)
    selected_periodkey=sel_periodkey.selectbox("PeriodKey",periodkey_list,help=sel_periodkey_help)
    selected_cat=sel_cat.selectbox("Category",df_local[df_local["PeriodKey"]==selected_periodkey]["Category"].unique(),help=sel_cat_help)
    selected_subcat=sel_subcat.selectbox("Subcategory",df_local[(df_local["PeriodKey"]==selected_periodkey) & (df_local["Category"]==selected_cat)]["Subcategory"].unique(),help=sel_subcat_help)
    selected_country=sel_country.selectbox("Country",df_local[(df_local["PeriodKey"]==selected_periodkey) & (df_local["Category"]==selected_cat) & (df_local["Subcategory"]==selected_subcat)]["Country"].unique(),help=sel_cty_help)

    df_local = df_local[(df_local["PeriodKey"] == selected_periodkey) & (df_local["Category"] == selected_cat) & (df_local["Subcategory"] == selected_subcat) & (df_local["Country"] == selected_country)]

    tab.markdown("**Step 2:** Choose Type/Subtype. *Only select Type/Subtype with similar Segments or Brands for comparison.*")
    sel_type,sel_subtype=tab.columns(2)
    # selected_type=sel_type.selectbox("Type",type_list,help=sel_type_help)
    # selected_subtype=sel_subtype.selectbox("Subtype",df_local[df_local["Type"]==selected_type]["Subtype"].unique(),help=sel_subtype_help)
    # selected_subtype=sel_subtype.multiselect("Subtype",df_local[df_local["Type"]==selected_type]["Subtype"].unique(),help=sel_subtype_help)
    selected_type=sel_type.multiselect("Type",type_list,help=sel_type_help)
    selected_subtype=sel_subtype.multiselect("Subtype",df_local[df_local["Type"].isin(selected_type)]["Subtype"].unique(),help=sel_subtype_help)

    df_local = df_local[(df_local["Type"].isin(selected_type)) & (df_local["Subtype"].isin(selected_subtype))]
    shortlist_segment_list = ['[Segment Comparison]'] + list(df_local['Segment'].unique())
    shortlist_brand_list = ['[Brand Comparison]'] + list(df_local['Brand'].unique())

    tab.markdown("**Step 3:** Choose either Segment or Brand comparison. Whichever is chosen, keep that selection as [Segment Comparison] or [Brand Comparison].")
    statsig_split,sel_segment,sel_brand = tab.columns(3)
    selected_split=statsig_split.selectbox("Segment or Brand Comparison",["Segment","Brand"],help=sel_split_help)
    if selected_split=="Segment":
        shortlist_brand_list.remove('[Brand Comparison]')
    elif selected_split=="Brand":
        shortlist_segment_list.remove('[Segment Comparison]')
    selected_segment = sel_segment.selectbox("Segment",shortlist_segment_list,help=sel_segment_help)
    selected_brand = sel_brand.selectbox("Brand",shortlist_brand_list,help=sel_brand_help)
    if selected_split == "Segment":
        df_local = df_local[(df_local['Brand']==selected_brand)]
    elif selected_split == "Brand":
        df_local = df_local[(df_local['Segment']==selected_segment)]
    df_local = df_local.sort_values(by=['Measure Value'],ascending=False)

    tab.markdown("**Step 4:** Choose type of stat-sig, and Benchmark Brand/Segment if applicable.")
    statsig_type, benchmark_target = tab.columns(2)
    selected_statsig_type=statsig_type.selectbox("Statsig Method",["Benchmark","Max","Benchmark (Inverse)"],help=sel_statsig_help)
    benchmark_target=benchmark_target.selectbox(f"Benchmark {selected_split}",get_benchmark(df_local,selected_split,selected_statsig_type),help=sel_benchmark_help)

    ## Find the unique list of split
    unique_split=get_benchmark(df_local,selected_split,"Benchmark")
    unique_split.sort()

    return df_local


def set_decimal_place_box(tab=st, key=str):
    """Implement drop-down box to select 0,1,2 decimal places."""
    dp_tooltip = "Choose desired number of decimal places. If 0 selected, will only return whole numbers."
    dp = tab.selectbox("Select desired decimal places.", options=range(0, 3), help=dp_tooltip, key=key)
    return dp


def dande_statsig_select(df, tab=st, segment=str):
    """Drivers and Equity Tab: Implement multiple filters drop-down boxes."""
    period_key_sel_help = "Select Period from available list of Period Keys."
    category_sel_help = "Select Category from available list of Segments."
    subcategory_sel_help = "Select SubCategory from available list of Segments."
    country_sel_help = "Select Country from available list of Segments."
    segment_sel_help = "Select Segment from available list of Segments."
    period_key_sel, category_sel, subcategory_sel, country_sel, segment_sel = tab.columns(5)
    period_key = period_key_sel.selectbox("Select PeriodKey", list(df['PeriodKey'].unique()), help=period_key_sel_help, key="PeriodKey_selection")
    category = category_sel.selectbox("Select Category", list(df['Category'].unique()), help=category_sel_help, key="Category_selection")
    subcategory = subcategory_sel.selectbox("Select SubCategory", list(df['SubCategory'].unique()), help=subcategory_sel_help, key="Subcategory_selection")
    country = country_sel.selectbox("Select Country", list(df['Country'].unique()), help=country_sel_help, key="Country_selection")
    segment = segment_sel.selectbox("Select Segment", list(df['Segment'].unique()), help=segment_sel_help, key="Segment_selection")

    bencmark_target, base_sel = tab.columns(2)
    selected_statsig_type = "Benchmark"
    sel_benchmark_help = "Select Brand to be benchmarked against."
    # sel_statsig_help="Select type of statsig, Benchmark or Max logic"
    # sel_benchmark_help="Select Benchmark Brand/Segment (Appicable to Statsig method = benchmark only)"
    # statsig_type,bencmark_target,base_sel=tab.columns(3)
    # selected_statsig_type=statsig_type.selectbox("Statsig Method",["Benchmark","Max"],help=sel_statsig_help,key=segment+"_type")

    if selected_statsig_type == "Max":
        df_selection = [None]
    elif selected_statsig_type == "Benchmark":
        df_selection = df[(df['PeriodKey'] == period_key) & (df["Segment"] == segment)]["Brand"].unique()
        df_selection = sorted(df_selection)
    benchmark_target = bencmark_target.selectbox("Benchmark", df_selection, help=sel_benchmark_help,
                                                 key=segment + "_target")

    base_tooltip = "Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""

    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}

    def format_func(option):
        return base_dict[option]

    base = base_sel.selectbox("Base (Overall Segment with Brand = All)", options=list(base_dict.keys()),
                              format_func=format_func, help=base_tooltip, key=segment)
    return period_key, category, subcategory, country, segment, selected_statsig_type, benchmark_target, base


def gen_output_xl(df,base,name=None,tab=None):
    """Implement Button to download results table as Excel file, with stat-sig styling."""
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False,na_rep="-")
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()
        if name==None:
            try:
                name=selected_type[0]+" - "+ selected_subtype[0] + "_"+ selected_split + "_"+ selected_statsig_type+".xlsx"
                # name=selected_cat[:10]+"_"+selected_subcat[:10]+"_"+selected_country[:3]+"_"+selected_type+" - "+ selected_subtype[0] + "_"+ selected_split + "_"+ selected_statsig_type+"_"+str(base) + ".xlsx"
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


def gen_output_pptx(df,decimal_place, slide_type="Performance",base=1200,name=None,tab=None,statsig=str,key=str, country="", segment="", brand=""):
    """Implement Button to call main_execute from barebones_ver2_ss.py, which downloads results table as PPT file."""
    df=df.dropna(axis=0, how='all')
    # df = df.dropna(axis=1, how='all')
    drop_cols_subset = [col for col in df.columns if col.startswith("Segment") or col.startswith("Period")]
    cols_to_drop = [col for col in drop_cols_subset if df[col].isna().all()]
    df = df.drop(columns=cols_to_drop)

    statsig=statsig.lower()
    buffer = main_execute(df,statsig,base,decimal_place=decimal_place,slide_type=slide_type, country=country, segment=segment, brand=brand)
    if name==None:
        name = country + "_" + segment + "_" + brand + "_" + statsig + "_output.pptx"
        name = name.lstrip("_")
        # try:
        #     name=selected_cat[:10]+"_"+selected_subcat[:10]+"_"+selected_country[:3]+"_"+selected_type[0]+" - "+ selected_subtype[0] + "_"+ selected_split + "_"+ selected_statsig_type+"_"+str(base) + ".pptx"
        # except:
        #     name="default.pptx"
    if tab==None:
        tab = st

    tab.download_button(
        label="Download PPT",
        data=buffer,
        file_name=name,
        key=key
    )


###### Stat-Sig ######
def find_largest_and_second_largest(numbers_list):
    """Get first and second largest based on values of each row for Max Stat-sig analysis"""
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


def find_threshold(number,base=None, statsig_type=""):
    """Get threshold value for segment / brand statistical significance comparison"""
    statsig_factor = 1.25
    if statsig_type == "PoP":
        statsig_factor = 1
    if ((type(number) == str) or (number is None)):
        return "Error"
    elif (base==None) or (base>1100):
        if number>45:
            return (4.1*statsig_factor)
        elif number >30:
            return (1.76*statsig_factor)
        elif number >15:
            return (0.98*statsig_factor)
        elif number > 0:
            return (0.56*statsig_factor)
        else:
            return "Error"
    elif (base>750):
        if number>45:
            return (4.7*statsig_factor)
        elif number >30:
            return (2.21*statsig_factor)
        elif number >15:
            return (1.73*statsig_factor)
        elif number > 0:
            return (0.97*statsig_factor)
        else:
            return "Error"
    elif (base>300):
        if number>45:
            return (6.3*statsig_factor)
        elif number >30:
            return (3.22*statsig_factor)
        elif number >15:
            return (2.2*statsig_factor)
        elif number > 0:
            return (1.23*statsig_factor)
        else:
            return "Error"
    else:
        return "Base less than 300!!"


def apply_statsig(df_pivot, statsig_type, base, first_col=1):
    """
        Max Stat-sig: Compare data points of cross segments to get diff between 1st/2nd largest and determine stat-significance of largest value. //
        Benchmark Stat-sig: Compare each data point to ref benchmark value to determine stat-superior/inferior of benchmark value against other value. //
        Set cell font colour to green/red if diff exceeds Stat-Sig threshold.
    """
    def max_logic(row, format, first_col):
        values = row[first_col:]
        highlight = None
        largest, sec_largest = find_largest_and_second_largest(values)

        threshold = find_threshold(largest, base)
        # st.write(row,largest,sec_largest,threshold)
        if largest is not None and sec_largest is not None:
            if largest - sec_largest > threshold:
                highlight = largest
        # elif largest is not None:
        #     highlight=largest

        return_list = []
        for col in row:
            try:
                if int(round(col, 0)) == int(round(highlight, 0)):
                    return_list.append(format)
                else:
                    return_list.append('')
            except:
                return_list.append('')

        return return_list
    def benchmark_logic(row, sup_format, inf_format, first_col):
        benchmark = row.iloc[first_col]
        values = row.iloc[first_col + 1:]
        return_list = [0] * (first_col + 1)  ## Indexes and benchmark column
        format_return_list = []
        for value in values:
            threshold = find_threshold(value, base)
            if threshold == "Error":
                return_list.append(0)
            elif ((type(value) == str) or (type(benchmark) == str)):  ## to handle when either value or benchmark is "-"
                return_list.append(0)
            elif value - benchmark > threshold:
                return_list.append(1)
            elif value - benchmark < -threshold:
                return_list.append(-1)
            else:
                return_list.append(0)

        for value in return_list:
            if value == 0:
                format_return_list.append("")
            elif value == 1:
                format_return_list.append(sup_format)
            elif value == -1:
                format_return_list.append(inf_format)

        return format_return_list
    def pop_logic(row, sup_format, inf_format, first_col):
        values = row[first_col:]
        return_list = [0] * (first_col)  ## Indexes and benchmark column
        format_return_list = []
        for idx, current_value in enumerate(values):
            previous_value = values.iloc[idx - 1]
            diff = ((current_value - previous_value) / previous_value) * 100
            threshold = find_threshold(previous_value, base, statsig_type="PoP")
            if threshold == "Error":
                return_list.append(0)
            elif ((type(current_value) == str) or (type(previous_value) == str)):  ## to handle when either value is "-"
                return_list.append(0)
            elif diff > threshold:
                return_list.append(1)
            elif diff < -threshold:
                return_list.append(-1)
            else:
                return_list.append(0)

        for value in return_list:
            if value == 0:
                format_return_list.append("")
            elif value == 1:
                format_return_list.append(inf_format)
            elif value == -1:
                format_return_list.append(sup_format)

        return format_return_list

    if statsig_type == "Max":
        df_pivot = df_pivot.style.apply(lambda x: max_logic(x, 'color: green;background-color:lightgreen', first_col),axis=1)
    elif statsig_type == "Benchmark":
        df_pivot = df_pivot.style.apply(lambda x: benchmark_logic(x, 'color: red;background-color:pink', 'color: green;background-color:lightgreen', first_col), axis=1)
    elif statsig_type == "Benchmark (Inverse)":
        df_pivot = df_pivot.style.apply(lambda x: benchmark_logic(x, 'color: green;background-color:lightgreen', 'color: red;background-color:pink', first_col), axis=1)
    elif statsig_type == "PoP":
        df_pivot = df_pivot.style.apply(lambda x: pop_logic(x, 'color: red;background-color:pink', 'color: green;background-color:lightgreen', first_col), axis=1)

    return df_pivot


def statsig_tab_highlight(updated_df,ss_type,base,decimal_place, tab=st):
    """Implement styling for statistically significant cells in results table."""
    #   Convert selected columns to float
    # columns_to_convert = updated_df.columns[1:]

    # updated_df[columns_to_convert] = updated_df[columns_to_convert].astype(float)
    updated_df = updated_df.dropna(axis=0, how='all')
    # df = df.dropna(axis=1, how='all')
    drop_cols_subset = [col for col in updated_df.columns if col.startswith("Segment") or col.startswith("Period")]
    cols_to_drop = [col for col in drop_cols_subset if updated_df[col].isna().all()]
    updated_df = updated_df.drop(columns=cols_to_drop)

    for column in updated_df.columns[1:]:
        updated_df[column] = pd.to_numeric(updated_df[column].str.replace('%', ''), errors='coerce')

    updated_df = apply_statsig(updated_df, ss_type, base, first_col=1)

#   Change from 1 dp to no dp
#     updated_df = updated_df.format(precision=1, na_rep='-')
#     updated_df = updated_df.format(precision=0, na_rep='-')
    updated_df = updated_df.format(precision=decimal_place, na_rep='-')
    return updated_df


###### Main Functions ######
def multi_select_and_df(df1, tab=st, decimal_place=1):
    """Performance Tab: Implement Base selection drop-down box, comparison Brands/Segments selection, results table, Download Buttons."""
    global selected_type, selected_subtype, selected_split, unique_split, selected_statsig_type, benchmark_target, selected_segment, selected_brand
    # global base,dataframe_show,df_pivot
    # split_list, dataframe_show = tab.columns([1, 3])
    split_list, dataframe_show = tab.columns([1, 2])

    ### Multi select
    i = 1
    shortlist_brand_segment = []
    if "Benchmark" in selected_statsig_type:
        shortlist_brand_segment = [benchmark_target]

    # split_list.subheader("Segment List")
    split_list.subheader("Segment/Brand List")

    base_tooltip = "Pls input No_of_People from Base sheet for the overall base selected (when Brand=All and Segment is largest possible segment group),e.g. Segment=\"Consumers using laxatives\" instead of Segment=\"Consumers using natural laxatives\""
    base_dict = {301: ">300", 751: "> 750", 1101: "> 1100"}

    def format_func(option):
        return base_dict[option]

    base = split_list.selectbox("Base", options=list(base_dict.keys()), format_func=format_func, help=base_tooltip)
    # base=split_list.number_input("Base",301,10000,1200,50,help=base_tooltip)
    # ## Select all checkbox WIP, not working as intended
    # all_brand_segment = list(unique_split).copy()
    # if ("Benchmark" in selected_statsig_type) & (benchmark_target==True):
    #     all_brand_segment.remove(benchmark_target)
    # if (len(selected_type)>=1) & (len(selected_subtype)>=1):
    #     check_all_checked = split_list.checkbox(label="[All]", value=all_brand_segment,key="checkall")
    #     if check_all_checked:
    #         shortlist_brand_segment.extend(all_brand_segment)
    for seg_brand_segment in unique_split:
        if ("Benchmark" in selected_statsig_type) & (seg_brand_segment == benchmark_target):
            continue
        seg_brand_checked = split_list.checkbox(seg_brand_segment, key="seg_brand_segment" + str(i))
        i = i + 1
        if seg_brand_checked:
            shortlist_brand_segment.append(seg_brand_segment)
    # shortlist_brand_segment = list(set(shortlist_brand_segment))

    # if "Benchmark" in selected_statsig_type:
    #
    #     for seg_brand_segment in unique_split:
    #         if seg_brand_segment!=benchmark_target:
    #             seg_brand_checked=split_list.checkbox(seg_brand_segment,key="seg_brand_segment"+str(i))
    #             i=i+1
    #             if seg_brand_checked:
    #                 shortlist_brand_segment.append(seg_brand_segment)
    # else:
    #     for seg_brand_segment in unique_split:
    #         seg_brand_checked=split_list.checkbox(seg_brand_segment,key="seg_brand_segment"+str(i))
    #         i=i+1
    #         if seg_brand_checked:
    #             shortlist_brand_segment.append(seg_brand_segment)

    ### Multi select
    dataframe_show.subheader("Statistical Significance table")
    ## Column list (Index)
    # col_list=list(df1.columns)
    # col_list.remove("Measure Value")
    # col_list.remove(selected_split)
    # # col_list.remove("No_of_People")
    # col_list.remove("Subtype")
    remove_list = ["Measure Value", selected_split]
    col_list = [col for col in df1.columns if col not in remove_list]

    ## Filter df
    if selected_split == "Brand":
        filtered_df = df1[(df1["Type"].isin(selected_type)) & (df1["Subtype"].isin(selected_subtype)) & (
            df1["Brand"].isin(shortlist_brand_segment))]
    elif selected_split == "Segment":
        filtered_df = df1[(df1["Type"].isin(selected_type)) & (df1["Subtype"].isin(selected_subtype)) & (
            df1["Segment"].isin(shortlist_brand_segment))]
    # if selected_statsig_type=="Max":
    #     if selected_split=="Brand":
    #         filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype)) & (df1["Brand"].isin(shortlist_brand_segment))]
    #     elif selected_split=="Segment":
    #         filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype)) & (df1["Segment"].isin(shortlist_brand_segment))]
    # elif "Benchmark" in selected_statsig_type:
    #     if selected_split=="Brand":
    #         filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype)) & (df1["Brand"].isin(shortlist_brand_segment+[benchmark_target]))]
    #     elif selected_split=="Segment":
    #         filtered_df=df1[(df1["Type"]==selected_type) & (df1["Subtype"].isin(selected_subtype)) & (df1["Segment"].isin(shortlist_brand_segment+[benchmark_target]))]

    ## Pivot
    df_pivot = pd.DataFrame()
    try:
        # df_pivot=filtered_df.pivot(index=index,columns=selected_split,values="Measure Value")
        # df_pivot=df_pivot.reset_index()
        # if selected_split == "Segment":
        #     col_list.remove("Brand")
        # elif selected_split == "Brand":
        #     col_list.remove("Segment")
        df_pivot = filtered_df.pivot(index=col_list, columns=selected_split, values="Measure Value")
        df_pivot = df_pivot.reset_index()
        # remove_list = ["PeriodKey","Country","Category","Subcategory","Segment","Brand"]
        # col_list = [col for col in col_list if col not in remove_list]
        keys_list = ['Type', 'Subtype', 'Content']
        df_cols_list = keys_list + shortlist_brand_segment
        df_pivot = df_pivot[df_cols_list]
        df_pivot.loc[df_pivot['Content'].isna(), 'Content'] = df_pivot['Type'].astype(str) + " " + df_pivot[
            'Subtype'].astype(str)
        df_pivot['Content'] = df_pivot['Content'].apply(lambda x: str(x)[:1].upper() + str(x)[1:])
        # try:
        #     df_pivot=df_pivot[col_list+shortlist_brand_segment] ## Sort columns
        # except KeyError:
        #     if (len(selected_type) < 1) | (len(selected_subtype) < 1):
        #         split_list.write("Please select 1 or more Type/Subtype")
        #     else:
        #         split_list.write("<span style='font-size:20px;padding-left: 5px;'> :exclamation: :exclamation:  Pivot Error :exclamation:  :exclamation: </span>", unsafe_allow_html=True)

        # if selected_statsig_type == "Max":
        #     if len(selected_subtype) < 1:
        #         split_list.write("Please select 1 or more subtype")
        #     else:
        #         try:
        #             df_pivot=df_pivot[col_list+shortlist_brand_segment] ## Sort columns
        #         except:
        #             split_list.write("<span style='font-size:20px;padding-left: 5px;'> :exclamation: :exclamation:  Pivot Error :exclamation:  :exclamation: </span>", unsafe_allow_html=True)
        #
        #
        # elif "Benchmark" in selected_statsig_type:
        #     try:
        #         df_pivot=df_pivot[col_list+[benchmark_target]+shortlist_brand_segment] ## Sort columns
        #     except:
        #
        #         if len(selected_subtype) >= 1:
        #             split_list.write("<span style='font-size:20px;padding-left: 5px;'> :exclamation: :exclamation:  Pivot Error :exclamation:  :exclamation: </span>", unsafe_allow_html=True)
        #         else:
        #              split_list.write("🙁 Please select 1 or more subtype")
        # df_pivot=df_pivot.drop(columns=["Country","Category","Subcategory","Type","Subtype"])
        ## SORTING
        df_cols_sort_list = df_cols_list.copy()
        df_cols_sort_list.remove('Content')
        sort_true_false_list = [True, True] + [False for _ in range(len(shortlist_brand_segment))]
        df_pivot = df_pivot.sort_values(df_cols_sort_list, ascending=sort_true_false_list, na_position="last")
        # df_pivot_cols = ['Type','Subtype','Content',benchmark_target] + shortlist_brand_segment
        # df_pivot = df_pivot[df_pivot_cols]

        styled_table_results = dataframe_show.container()
        decimal_place_placeholder = dataframe_show.container()
        with decimal_place_placeholder:
            decimal_place = set_decimal_place_box(key="Performance_decimal_place_box")
        df_pivot_styler = apply_statsig(df_pivot, selected_statsig_type, base, first_col=3)
        df_pivot_styler_formatted = df_pivot_styler.format(precision=decimal_place, na_rep='-')

        ## To fill na with "-"
        # df_pivot = df_pivot.fillna("")

        # df_pivot = df_pivot.fillna(None)
        ## END To fill na with "-"
        with styled_table_results:
            st.dataframe(df_pivot_styler_formatted, hide_index=True)

        # dataframe_show.write("Error")
        ## To fill "" with "-" for output
        # df_pivot = df_pivot.applymap(lambda x: st.write(x))

        gen_output_pptx(df_pivot, decimal_place=decimal_place, statsig=selected_statsig_type, base=base,
                        tab=dataframe_show, key="Performance" + selected_statsig_type,
                        country=selected_country, segment=selected_segment, brand=selected_brand)
        gen_output_xl(df_pivot_styler_formatted, base, tab=dataframe_show, name=f"Performance_Statsig_output.xlsx")

    # except TypeError:
    #     dataframe_show.write("Multiple Type and Subtype columns")

    except Exception as Esc:
        # dataframe_show.write(Esc)
        # if (df_pivot.empty):
        if (len(selected_type) < 1) | (len(selected_subtype) < 1):
            dataframe_show.write("Please select 1 or more Type/Subtype")
        elif ((df_pivot.empty) & ((len(selected_type) > 0) | (len(selected_subtype) > 0))) | (
                (not df_pivot.empty) & (len(shortlist_brand_segment) <= 1)):
            ## WIP: Error messaging not working as intended.
            if len(unique_split) <= 1:
                dataframe_show.write("Under Step 3, pls select appropriate option of Brand or Segment for comparison.")
            elif (len(unique_split) > 1) & (selected_statsig_type == "Max"):
                dataframe_show.write("Pls select at least 2 Brands/Segments for comparison.")
            elif (len(unique_split) > 1) & ("Benchmark" in selected_statsig_type):
                dataframe_show.write("Pls select at least 1 Brand/Segment for comparison, in addition to Benchmark.")
        elif (not df_pivot.empty) & (len(shortlist_brand_segment) > 1):
            # elif (not df_pivot.empty) & (((len(shortlist_brand_segment)>1)&(selected_statsig_type=="Benchmark"))|((len(shortlist_brand_segment)>0)&(selected_statsig_type=="Max"))):
            dataframe_show.write(
                "<span style='font-size:20px;padding-left: 10px;'> :exclamation: :exclamation:  DATA ERROR :exclamation:  :exclamation: </span>",
                unsafe_allow_html=True)
            dataframe_show.write(Esc)
        else:
            dataframe_show.write("🙁 Please contact DS for help.")


def dande_tab_execute(df,tab=st):
    """Drivers and Equity Tab: Implement results table, Download Buttons."""
    ## Select box for statsig method and benchmarl
    period_key, category, subcategory, country, segment, dande_selected_statsig,dande_benchmark,base=dande_statsig_select(df,tab)

    ## Get CBI
    df_cbi=df.copy()
    df_cbi=df_cbi[(df_cbi["PeriodKey"]==period_key)&(df_cbi["Category"]==category)&(df_cbi["SubCategory"]==subcategory)&(df_cbi["Country"]==country)&(df_cbi["Segment"]==segment)]
    # df_cbi=df_cbi[(df_cbi["PeriodKey"]==period_key)&(df_cbi["Segment"]==segment)]
    df_cbi=df_cbi.drop(columns=["Driver_Score","Driver","Equity_Score"])
    df_cbi=df_cbi.drop_duplicates()
    df_cbi["Driver"] = "CBI"
    df_cbi=df_cbi.rename(columns={"CBI":"Equity_Score"})
    # df_with_cbi = pd.concat([df[df["Segment"]==segment],df_cbi])
    df_equity = df[(df["PeriodKey"]==period_key)&(df["Segment"]==segment)]
    df_equity = df_equity.drop(columns=["CBI"])
    df_with_cbi = pd.concat((df_equity, df_cbi), axis=0).reset_index(drop=True)


    ## Pivot
    df_pivot=df_with_cbi[(df_with_cbi["PeriodKey"]==period_key)&(df_with_cbi["Segment"]==segment)].pivot(index=["Driver_Score","Driver"],columns="Brand",values="Equity_Score").reset_index()
    df_pivot=df_pivot.sort_values("Driver_Score",ascending=False).reset_index(drop=True)

    ## Rearrange column if is benchamrk
    if dande_selected_statsig == "Benchmark":
        other_brand=df[(df["PeriodKey"]==period_key)&(df["Segment"]==segment)]["Brand"].unique().tolist()
        other_brand.remove(dande_benchmark)
        # other_brand=other_brand.to_list()
        # tab.dataframe(other_brand)

        df_pivot=df_pivot[["Driver_Score","Driver",dande_benchmark]+other_brand]


    df_pivot = df_pivot.convert_dtypes()
    updated_df=apply_statsig(df_pivot,dande_selected_statsig,base,first_col=2)

    tab.dataframe(updated_df, hide_index=True)
    # decimal_place = set_decimal_place_box(key="Drivers_decimal_place_box")
    gen_output_pptx(df_pivot,decimal_place=0, statsig=dande_selected_statsig,slide_type="Drivers",base=base,name="DriversEquity_"+segment+".pptx",tab=tab,key="DriversEquity", segment=segment)
    gen_output_xl(updated_df,base,name="DriversEquity_Statsig_output.xlsx",tab=tab)
    # gen_output_xl(updated_df,base,name="DriversEquity_"+segment+".xlsx",tab=tab)


def app():
    """Implement 4 tabs with respective filters, result tables, downloads."""
    st.title('Self service Statistical Significance app')
    # allow_cors()

    global buffer
    buffer = io.BytesIO()
    statsig_tab, pop_tab, dande_tab,fixed_tab = st.tabs(["Statistical Significant", 'PoP Statistical Significant', "Drivers and Equity", "Performance", ])
    # fixed_tab, statsig_tab, pop_tab, dande_tab = st.tabs([ "Performance", "Statistical Significant", 'PoP Statistical Significant', "Drivers and Equity"])

    ## Statsig _tab
    statsig_tab.write("If **Benchmark** is the selected statsig method, **first column** would be assumed to be the benchmark value")
    updated_df = empty_df(statsig_tab)
    ss_type, base = statsig_tab_sel_box(statsig_tab)
    statsig_placeholder = statsig_tab.container()
    decimal_place = set_decimal_place_box(statsig_tab, key="decimal_place_box")
    applied_df = statsig_tab_highlight(updated_df, ss_type, base, decimal_place, statsig_tab)
    with statsig_placeholder:
        st.subheader("Results")
        st.dataframe(applied_df)
    gen_output_pptx(updated_df, decimal_place=decimal_place, statsig=ss_type, base=base, tab=statsig_tab,key="self_service_within_period")
    gen_output_xl(applied_df, base, tab=statsig_tab, name=ss_type+"_Statsig_output.xlsx")
    ## PoP tab
    pop_tab.write("Input **Content** in the first column, and available periods data in chronological order.")
    updated_pop_df = empty_pop_df(pop_tab)
    pop_tab.write("*Select PoP based on previous period of comparison, as Base may change over time.*")
    pop_base = pop_statsig_tab_sel_box(pop_tab)
    pop_placeholder = pop_tab.container()
    pop_decimal_place = set_decimal_place_box(pop_tab, key="PoP_decimal_place_box")
    applied_pop_df = statsig_tab_highlight(updated_pop_df, "PoP", pop_base, pop_decimal_place, pop_tab)
    with pop_placeholder:
        st.subheader("Results")
        st.dataframe(applied_pop_df)
    # gen_output_xl(applied_pop_df, base, tab=pop_tab, name="PoP_Statsig_output.xlsx")  ## COMMENTED OUT, CAUSES OTHER XL OUTPUT TO BE CORRUPTED


    # ## D&E _tab
    df2,f2=get_file_input(dande_tab,sheet_name="Drivers")
    if f2 is not None:
        # dande_segment(df2,dande_tab)
        dande_tab_execute(df2,dande_tab)


    # Performance tab
    df1,f1=get_file_input(fixed_tab,sheet_name="Performance")
    if f1 is not None:
        global split_list
        # split_list=["Segment","Brand"]
        get_unique(df1)
        df1=performance_select_box(df1,fixed_tab)
        multi_select_and_df(df1,fixed_tab)

    
    # if (f1 is None):
    #     return

if __name__ == '__main__':
    st.set_page_config(page_title="Self-service statsig", page_icon="📊", layout="wide")
    # st.set_page_config(page_title="Self Service statsig", page_icon=":chart_with_upwards_trend:", layout="wide")

    app()
    # print("!!APP REFRESHED!!")
