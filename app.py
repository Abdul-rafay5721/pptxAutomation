from flask import Flask, render_template, request, redirect, url_for, send_file, after_this_request
import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt  # Add this import
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor  # Add this import
from werkzeug.utils import secure_filename

app = Flask(__name__)
# Configure upload folder
app.config['UPLOAD_FOLDER'] = '/tmp'  # Use /tmp directory on Linux
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limit file size to 16MB

# Function to replace placeholders in the presentation
def replace_placeholders(presentation, data):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        for key, value in data.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in text:
                                text = text.replace(placeholder, value)
                        run.text = text

# Function to insert numbered bullet points into a specific text frame
def insert_bullet_points(slide, placeholder, bullet_points):
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, "")
                        for i, bullet_point in enumerate(bullet_points, start=1):
                            p = text_frame.add_paragraph()
                            p.text = bullet_point
                            p.level = 0
                            p.bullet = False  # Ensure the paragraph is not a bullet point
                            p.number = i  # Set the numbering

# Function to insert deliverable bullets without amounts into a specific text frame or table cell
def insert_deliverable_bullets(slide, placeholder, deliverable_bullets):
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        cell.text = ""  # Clear the cell text completely
                        for i, deliverable in enumerate(deliverable_bullets, start=1):
                            p = cell.text_frame.add_paragraph()
                            p.text = f"{i}. {deliverable}"
                            p.level = 0  # Ensure level 0 for standard bullets
                            p.space_before = Pt(0)  # Remove space before paragraph
                            p.space_after = Pt(0)   # Remove space after paragraph
                        return  # Exit after replacing the placeholder
        elif shape.has_text_frame:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = ""  # Clear the text completely
                        for i, deliverable in enumerate(deliverable_bullets, start=1):
                            p = text_frame.add_paragraph()
                            p.text = f"{i}. {deliverable}"
                            p.level = 0  # Standard bullet level
                            p.space_before = Pt(0)  # Remove space before paragraph
                            p.space_after = Pt(0)   # Remove space after paragraph
                        return  # Exit after replacing the placeholder

# Function to insert deliverables and amounts into specific placeholders within a table
def insert_deliverables_and_amounts(presentation, deliverables, amounts):
    placeholders = [f"{{{{deliverable{i}}}}}" for i in range(1, 6)]
    amount_placeholders = [f"{{{{amount{i}}}}}" for i in range(1, 6)]
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                deliverable_rows = []
                for i, row in enumerate(table.rows):
                    if i == 0:  # Skip header row
                        continue
                    if any(placeholders[j] in cell.text for j in range(len(placeholders)) for cell in row.cells):
                        deliverable_rows.append(row)
                for i, row in enumerate(deliverable_rows):
                    if i < len(deliverables):
                        for cell in row.cells:
                            if placeholders[i] in cell.text:
                                cell.text = cell.text.replace(placeholders[i], deliverables[i])
                            if amount_placeholders[i] in cell.text:
                                cell.text = cell.text.replace(amount_placeholders[i], amounts[i])
                    else:
                        table._tbl.remove(row._tr)  # Remove extra deliverable rows

# Function to replace an image placeholder in the presentation
def replace_image_placeholder(presentation, placeholder, image_path):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, "")
                            left = shape.left
                            top = shape.top
                            slide.shapes.add_picture(image_path, left, top, width=Inches(1.75), height=Inches(1.0))
                            return

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    # Get template file
    if 'template' not in request.files:
        return 'No template file uploaded', 400
    template_file = request.files['template']
    if template_file.filename == '':
        return 'No template selected', 400

    # Save template file temporarily with secure filename
    template_filename = secure_filename(template_file.filename)
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{template_filename}")
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], "output_proposal.pptx")
    template_file.save(template_path)

    data = {
        "company": request.form['company'],
        "Date": request.form['date'],
        "Service": request.form['service'],
        "activityPoints": request.form.getlist('activityPoints'),
        "companyNamePlaceholder": "{{companyName}}",
        "scopePlaceholder": "{{scope}}",
        "bulletPoints": request.form.getlist('bulletPoints'),
        "deliverableBullets": request.form.getlist('deliverableBullets'),
        "amounts": request.form.getlist('amounts'),
        "withoutVat": request.form['withoutVat'],
        "total": request.form['total']
    }

    presentation = Presentation(template_path)

    # Replace placeholders with data
    replace_placeholders(presentation, data)

    # Insert numbered bullet points where the scope placeholder is found
    for slide in presentation.slides:
        insert_bullet_points(slide, data["scopePlaceholder"], data["bulletPoints"])

    # Insert deliverable bullets where the deliverable placeholder is found
    for slide in presentation.slides:
        insert_deliverable_bullets(slide, "{{deliverable}}", data["deliverableBullets"])
        # Insert activities using the same function but with activity placeholder
        insert_deliverable_bullets(slide, "{{activity}}", data["activityPoints"])

    # Insert deliverables and amounts into specific placeholders within a table
    insert_deliverables_and_amounts(presentation, data["deliverableBullets"], data["amounts"])

    # Replace placeholders with data including totals
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        if "{{withoutVat}}" in cell.text:
                            cell.text = cell.text.replace("{{withoutVat}}", data["withoutVat"])
                            # Set text color to white
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(255, 255, 255)
                        if "{{total}}" in cell.text:
                            cell.text = cell.text.replace("{{total}}", data["total"])
                            # Set text color to white
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(255, 255, 255)

    # Replace image placeholder with the uploaded image
    if 'image' in request.files:
        image_file = request.files['image']
        if image_file.filename != '':
            image_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(image_file.filename))
            image_file.save(image_path)
            replace_image_placeholder(presentation, "{{image}}", image_path)

    # Check if the output file exists and remove it
    if os.path.exists(output_path):
        os.remove(output_path)

    # Save the modified presentation
    presentation.save(output_path)

    # Clean up files immediately after sending
    @after_this_request
    def cleanup(response):
        try:
            if os.path.exists(template_path):
                os.remove(template_path)
            if os.path.exists(output_path):
                os.remove(output_path)
            if 'image' in request.files and request.files['image'].filename != '':
                image_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(request.files['image'].filename))
                if os.path.exists(image_path):
                    os.remove(image_path)
        except Exception as e:
            app.logger.error(f"Error cleaning up files: {e}")
        return response

    # Return file for download
    return send_file(
        output_path,
        as_attachment=True,
        download_name=f"proposal_{secure_filename(data['company'])}.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    # Production settings
    app.run(host='0.0.0.0', port=8080)
