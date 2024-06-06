from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
from myutils import pdf_bundle, add_page_numbers, create_page_pdf
import os, glob


tmploc= 'C:\\Users\\U059611\\flaskproject1\\temporary_files' 

def delete_temporary_files():
    temp_dir = tmploc  # Replace with the actual temporary directory path
    files_to_delete = glob.glob(os.path.join(temp_dir, '*'))
    for file_path in files_to_delete:
        os.remove(file_path)

delete_temporary_files()

app = Flask(__name__)
@app.route("/")
def main_func():
    return render_template('home.html')

@app.route('/bundle', methods=['POST'])
def bundle_pdf():
    # Get the uploaded PDF files and the desired name for the bundled PDF
    pdf_files = request.files.getlist('pdfFiles')
    bundled_pdf_name = request.form['bundledPdfName']
    num_pdf_files = int(request.form['numPdfFiles'])  # Retrieve the number of selected PDF files

    add_paging = request.form.get('addPaging')

    # Save the uploaded PDF files to a temporary directory
    temp_dir = tmploc  # Replace with the actual temporary directory path
    os.makedirs(temp_dir, exist_ok=True)
    saved_pdf_files = []
    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(temp_dir, pdf_file.filename)
        pdf_file.save(pdf_file_path)
        saved_pdf_files.append(pdf_file_path)

    # Call the pdf_bundle function with the list of saved PDF files
    output_pdf = os.path.join(temp_dir, bundled_pdf_name)

    pdf_bundle(saved_pdf_files, output_pdf)  
    

    if add_paging =='true':
        print('Paging!')
        print(add_paging)
        add_page_numbers(output_pdf)  
        
    # Return the filename of the bundled PDF
    return jsonify(filename=bundled_pdf_name)


@app.route('/download/<path:filename>', methods=['GET', 'POST'])
def download(filename):
    # Ensure the file exists.
    temp_dir = tmploc
    if not os.path.exists(os.path.join(temp_dir, filename)):
        return "File not found."

    #return send_from_directory(tmploc, filename)
    response = send_file(os.path.join(temp_dir, filename), as_attachment=True)

    return response




delete_temporary_files()
