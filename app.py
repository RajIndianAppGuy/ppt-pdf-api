from flask import Flask, request, send_file, jsonify, after_this_request
import os
import comtypes.client
import pythoncom
import threading
import time
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def ppt_to_pdf(ppt_file, pdf_file):
    powerpoint = None
    presentation = None

    try:
        pythoncom.CoInitialize()
        # Create a PowerPoint application object
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        # Open the PowerPoint file
        presentation = powerpoint.Presentations.Open(ppt_file)

        # Save as PDF
        presentation.SaveAs(pdf_file, FileFormat=32)  # 32 is the format for PDF

    except Exception as e:
        print(f"An error occurred: {e}")
        raise  # Re-raise the exception after logging it

    finally:
        # Ensure that the presentation and PowerPoint application are properly closed
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()
def delayed_cleanup(ppt_path, pdf_path, delay):
    time.sleep(delay)
    try:
        os.remove(ppt_path)
        os.remove(pdf_path)
    except Exception as e:
        print(f"Error deleting files: {e}")

@app.route('/', methods=['POST'])
def convert_ppt_to_pdf():
   
    ppt_file = request.files['file']
    ppt_filename = ppt_file.filename
    ppt_path = os.path.join(os.getcwd(),'uploads', ppt_filename)

    # Save the uploaded PowerPoint file
    ppt_file.save(ppt_path)

    # Define the output PDF file path
    pdf_filename = os.path.splitext(ppt_filename)[0] + '.pdf'
    pdf_path = os.path.join(os.getcwd(),'uploads', pdf_filename)
    print(ppt_path,pdf_path)
    try:
        # Convert the PowerPoint to PDF
        ppt_to_pdf(ppt_path, pdf_path)

        # Register a function to clean up files after response is sent
        @after_this_request
        def cleanup(response):
            threading.Thread(target=delayed_cleanup, args=(ppt_path, pdf_path, 10)).start()
            return response

        # Send the PDF file as a downloadable response
        return send_file(pdf_path, as_attachment=True)

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
