from flask import Flask, request, send_file, jsonify, after_this_request
import os
import subprocess
import threading
import time
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def ppt_to_pdf(ppt_file, pdf_file):
    try:
        # Use LibreOffice in headless mode to convert PPT to PDF
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', ppt_file, '--outdir', os.path.dirname(pdf_file)],
                       check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")
        raise  # Re-raise the exception after logging it

def delayed_cleanup(ppt_path, pdf_path, delay):
    time.sleep(delay)
    try:
        if os.path.exists(ppt_path):
            os.remove(ppt_path)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
    except Exception as e:
        print(f"Error deleting files: {e}")

@app.route('/', methods=['POST'])
def convert_ppt_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    ppt_file = request.files['file']
    ppt_filename = ppt_file.filename
    ppt_path = os.path.join(os.getcwd(), 'uploads', ppt_filename)

    # Save the uploaded PowerPoint file
    ppt_file.save(ppt_path)

    # Define the output PDF file path
    pdf_filename = os.path.splitext(ppt_filename)[0] + '.pdf'
    pdf_path = os.path.join(os.getcwd(), 'uploads', pdf_filename)

    try:
        # Convert the PowerPoint to PDF
        ppt_to_pdf(ppt_path, pdf_path)

        # Register a function to clean up files after response is sent
        @after_this_request
        def cleanup(response):
            # Start a thread to handle cleanup after a delay
            threading.Thread(target=delayed_cleanup, args=(ppt_path, pdf_path, 10)).start()
            return response

        # Send the PDF file as a downloadable response
        return send_file(pdf_path, as_attachment=True, attachment_filename=pdf_filename)

    except Exception as e:
        # Handle exceptions and cleanup if needed
        if os.path.exists(ppt_path):
            os.remove(ppt_path)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
