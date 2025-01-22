import pandas as pd
from docx import Document
import os
from docx import Document
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route("/")
def login():
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def handle_login():
    username = request.form.get("username")
    return redirect(url_for("details", username=username))

@app.route("/details")
def details():
    username = request.args.get("username")
    return render_template("details.html", username=username)

@app.route("/save_details", methods=["POST"])
def save_details():
    program = request.form.get("program")
    section = request.form.get("section")
    year = request.form.get("year")
    academic_year = request.form.get("academicYear")
    semester = request.form.get("semester")
    batch = request.form.get("batch")
    date = request.form.get("date")

    # Here you could store the data, e.g., in a database or a file.
    doc = Document("template.docx")
    placeholders = {
        "{{program}}": program,
        "{{sec}}": section,
        "{{year}}": year,
        "{{acyear}}": academic_year,
        "{{sem}}": semester,
        "{{batch}}": batch,
        "{{date}}": date
    }

    for placeholder, value in placeholders.items():
        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    doc.save("output.docx")

    return redirect(url_for("upload"))

@app.route("/upload")
def upload():
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    return render_template("upload.html")

@app.route("/upload_file", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return "No file part in the form."
#file storage
    file = request.files['file']
    
    if file.filename == '':
        return "No selected file."

    # Save the file
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        return f"File '{file.filename}' uploaded successfully!"
    
    return redirect(url_for("upload"))



if __name__ == "__main__":
    app.run(debug=True)

datafile=pd.ExcelFile("Book.xlsx")
subject=["MA3354","CS3301","CS3351","CS3352","CS3391"]
doc = Document("finaldoc.docx")
j=0
num=0
semester=4
samp=[]
def table1(sub,n):
    global samp
    global num
    global j
    df=pd.read_excel("Book.xlsx",sheet_name=sub)
    dic1=[]
    dic1.append(n)
    dic1.append(sub)
    dic1+=[sub]
    dic1+=[1]
    dic1+=[2]
    dic1+=[3]
    dic1+=[4]
    dic1+=["dfgbn"]
    dic1+=["fds"]
    dic1+=[len(df["name"])]
    dic1+=[len(df[df["UR"]!="ab"])]
    dic1+=[dic1[9]-dic1[10]]
    dic1+=[len(df[(df["UR"] != 'ab') & (df["UR"] != 'RA')])]
    dic1+=[dic1[10]-dic1[12]]
    dic1+=[f"{(dic1[12]/dic1[10])*100:.1f}"]
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
        num=len(dfs[dfs["ARH"]!=0])
    j=1
    dic1+=[f"{len(samp):.2f}"]
    dic1+=[int((len(samp)/len(df["name"]))*100)]#f"{(len(samp)/len(df["name"]))*100:.2f}"
    dic1+=[num]
    print(dic1)
    return dic1
def table2(sub,n):
    df=pd.read_excel("Book.xlsx",sheet_name=sub)
    dic1=[]
    dic1.append(sub)
    dic1+=[sub]
    dic1+=[1]
    dic1+=[2]
    dic1+=[3]
    dic1+=[4]
    dic1+=["dfgbn"]
    dic1+=["fds"]
    dic1+=[len(df["name"]),len(df["name"]),len(df["name"])]*3
    exm=["IA1","IA2","UR"]
    for i in range(2):
        dic1.append(len(df[(df[exm[i]]!="ab") & (df[exm[i]]!=0)]))
    for i in range(3):
        dic1.append(len(df["name"])-len(df[(df[exm[i]]!="ab") & (df[exm[i]]!=0)]))
    for i in range(2):
        dic1.append(len(df[(df[exm[i]] != 0) & (df[exm[i]]>35)]))
    dic1.append(len(df[(df["UR"] != 'ab') & (df["UR"] != 'RA')]))
    for i in range(2):
        dic1.append(len(df[(df[exm[i]] == 0) & (df[exm[i]]<35)]))
    dic1.append(len(df[(df["UR"] == 'ab') & (df["UR"] == 'RA')]))
    for i in range(2):
        dic1.append((len(df[(df[exm[i]] != 0) & (df[exm[i]]>35)])/len(df["name"]))*100)
    dic1.append((len(df[(df["UR"] != 'ab') & (df["UR"] != 'RA')])/len(df["name"]))*100)
    print(dic1)
    return dic1
'''def table3():
    """SEMESTER WISE ARREAR COUNT (SEMESTER WISE FROM SEM I TO CURRENT SEMESTER)"""
    df=pd.read_excel("Book.xlsx",sheet_name="Sheet1")
    lis=[]
    lis.append()'''



for table_index, table in enumerate(doc.tables): 
    if len(table.rows)>2:
        if len(table.rows) < len(subject):
            print(len(table.rows),len(subject))
            for _ in range(len(subject) - (len(table.rows)-2)):
                table.add_row()
        i=0
        for row_index, row in enumerate(table.rows):
            if row_index>=2:
                if table_index==0:
                    for col_in in range(15, 18):
                        start_cell = table.cell(2, col_in) 
                        for row_in in range(2, len(table.rows)): 
                            start_cell.merge(table.cell(row_in, col_in))
                    if (i<len(subject)):
                        data=table1(subject[i],i)
                elif table_index==2:
                    if (i<len(subject)):
                        data=table2(subject[i],i)
                elif table_index==4:
                    pass
                else:
                    continue 
                for cell_index, cell in enumerate(row.cells): 
                    cell.text = str(data[cell_index])
                i+=1
        i=0
    if table_index==4:
        break
doc.save("output_checked_cells.docx")