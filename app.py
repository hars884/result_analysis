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
