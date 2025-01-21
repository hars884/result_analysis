import pandas as pd
from docx import Document
datafile=pd.ExcelFile("Book.xlsx")
subject=["MA3354","CS3301","CS3351","CS3352","CS3391"]
doc = Document("finaldoc.docx")
i=0
j=0
num=0
samp=[]
def table1(sub):
    global samp
    global num
    global j
    df=pd.read_excel("Book.xlsx",sheet_name=sub)
    dic1={}
    dic1["sn"]=j
    dic1["couc"]=sub
    dic1["cout"]=sub
    dic1["l"]=1
    dic1["t"]=2
    dic1["p"]=3
    dic1["c"]=4
    dic1["couf"]="sdfgbn"
    dic1["dept"]="fds"
    dic1["tns"]=len(df["name"])
    dic1["nosa"]=len(df[df["UR"]!="ab"])
    dic1["nosab"]=dic1["tns"]-dic1["nosa"]
    dic1["nosp"] = len(df[(df["UR"] != 'ab') & (df["UR"] != 'RA')])
    dic1["nosf"]=dic1["nosa"]-dic1["nosp"]
    dic1["pp"]=f"{(dic1["nosp"]/dic1["nosa"])*100:.1f}"
    if j==0:
        oac=[]
        for sheet_name in datafile.sheet_names:
            dfs = pd.read_excel("Book.xlsx", sheet_name=sheet_name) 
            if 'UR' in dfs.columns and 'name' in dfs.columns:
                filtered_names = dfs[(dfs['UR'] != 'ab') & (dfs['UR'] != 'RA')]['name'].tolist()
            oac.append(filtered_names)
        samp=oac[0]
        for sl in range(1,len(oac)):
            samp1=[item for item in samp if item in oac[sl]]
            samp=samp1
        dfs = pd.read_excel("Book.xlsx", sheet_name="Sheet1")
        num=len(dfs[dfs["ARH"]=="ar"])
        print(samp," ",num)
    j=1
    dic1["cwa"]=f"{len(samp):.2f}"
    dic1["cwap"]=int((len(samp)/len(df["name"]))*100)#f"{(len(samp)/len(df["name"]))*100:.2f}"
    dic1["ywac"]=num
    return dic1
def table2(sub):
    dic1={}
    dic1["sn"]=j
    dic1["couc"]=sub
    dic1["cout"]=sub
    dic1["l"]=1
    dic1["t"]=2
    dic1["p"]=3
    dic1["c"]=4
    dic1["couf"]="sdfgbn"
    dic1["dept"]="fds"
for table_index, table in enumerate(doc.tables):
    #print('t',table_index,end=" ")
    #print('rc',len(table.rows),end=" ")
    table = doc.tables[0]
    if len(table.rows)>2:
        if len(table.rows) < len(subject):
            print(len(table.rows),len(subject))
            for _ in range(len(subject) - (len(table.rows)-2)):
                table.add_row()
        for col_index in range(15, 18):
            start_cell = table.cell(2, col_index) 
            for row_index in range(2, len(table.rows)): 
                start_cell.merge(table.cell(row_index, col_index))
        for row_index, row in enumerate(table.rows):
            #print('r',row_index,end=" ")
            #print('cc',len(row.cells),end=" ")
            data=table1(subject[i])
            da=list(data.keys())
            if row_index>=2:
                for cell_index, cell in enumerate(row.cells):
                    key = da[cell_index]  
                    cell.text = str(data[key])
                i+=1
    break
doc.save("output_checked_cells.docx")
print("I am harshini")
