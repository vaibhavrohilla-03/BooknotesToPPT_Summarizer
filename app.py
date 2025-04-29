import os
import tempfile
from flask import Flask, request, render_template, send_file, url_for, redirect, flash
import uuid
from werkzeug.utils import secure_filename
from datetime import datetime

from manual_topics_slide_generator import ManualTopicSlideGenerator

app = Flask(__name__)
app.secret_key = 'your_secret_key'  


UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
ALLOWED_EXTENSIONS = {'pdf'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check if textbook file is present
    if 'textbook' not in request.files:
        flash('Textbook file is required')
        return redirect(url_for('index'))
    
    textbook_file = request.files['textbook']
    
    # Check if textbook filename is not empty
    if textbook_file.filename == '':
        flash('No textbook file selected for uploading')
        return redirect(url_for('index'))
    
    # Check if file extension is allowed
    if not allowed_file(textbook_file.filename):
        flash('Only PDF files are allowed')
        return redirect(url_for('index'))
    
    manual_topics = request.form.get('manual_topics', '')
    if not manual_topics.strip():
        flash('Please enter at least one topic')
        return redirect(url_for('index'))
    
    job_id = str(uuid.uuid4())
    
    textbook_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_textbook.pdf")
    textbook_file.save(textbook_path)
    
    try:
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_slides.pptx")
        
        generator = ManualTopicSlideGenerator(textbook_path=textbook_path)
        
        generator.set_topics(manual_topics)
        
        # Run the pipeline
        generated_file = generator.run_pipeline(output_path=output_path)
        
        if generated_file:
            # Return success page with download link
            return render_template('success.html', 
                                job_id=job_id,
                                textbook_name=secure_filename(textbook_file.filename))
        else:
            flash('Error generating slides. Please check the console for details.')
            return redirect(url_for('index'))
    
    except Exception as e:
        # Handle errors
        flash(f'Error processing files: {str(e)}')
        if os.path.exists(textbook_path):
            os.remove(textbook_path)
        return redirect(url_for('index'))

@app.route('/download/<job_id>', methods=['GET'])
def download_file(job_id):
    # Security check: ensure job_id doesn't contain path traversal
    if '..' in job_id or '/' in job_id:
        flash('Invalid job ID')
        return redirect(url_for('index'))
    
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_slides.pptx")
    
    if not os.path.exists(output_path):
        flash('File not found or has been deleted')
        return redirect(url_for('index'))
    
    return send_file(output_path, as_attachment=True, download_name='generated_slides.pptx')

@app.route('/clean/<job_id>', methods=['GET'])
def clean_files(job_id):
    # Security check: ensure job_id doesn't contain path traversal
    if '..' in job_id or '/' in job_id:
        flash('Invalid job ID')
        return redirect(url_for('index'))
    
    textbook_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_textbook.pdf")
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_slides.pptx")
    
    for path in [textbook_path, output_path]:
        if os.path.exists(path):
            os.remove(path)
    
    flash('Files have been cleaned up')
    return redirect(url_for('index'))

@app.context_processor
def inject_current_year():
    return {'current_year': datetime.now().year}

if __name__ == '__main__':
    app.run(debug=True)