from flask import Flask, request, jsonify, send_file
from datetime import datetime
import activity
from docx import Document
from io import BytesIO

app = Flask(__name__)

@app.route('/process_data', methods=['POST'])
def process_data():
    try:
        # Record request receive time
        request_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        
        # Fill in the data
        process = activity.Process(request, request_time)
 
        # Process the text data to test if request received by API
        response_data = process.process_test()
        
        # TODO: Select appropriate template
        template_id = "004"

        # Copy to self.documents_folder name starting with result_{request_time}.docx
        process.copy_template(template_id)

        # Read result_{request_time}.docx
        # Over-write
        docx_path = process.overwrite_data()
        print("Path to document is:" + docx_path)
        # TODO: Return document as response_data
        # return jsonify({'response': 'Success'})        
        return send_file(docx_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': 'An error occurred.', 'details': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)