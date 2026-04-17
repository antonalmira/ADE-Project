from document_generator import DocGenerator
from chart_extractor import save_chart_screenshots
from utils import log_message

def generate_document(app):
    """Generate a new document using the selected data."""
    log_message("Starting document generation")
    # Extract chart screenshots before generating the document
    save_chart_screenshots(app)
    generator = DocGenerator(app, app.generated_document_path.text())
    generator.generate()
    log_message("Document generation completed")

def update_document(app):
    """Update an existing document with the selected data."""
    log_message("Starting document update")
    # Extract chart screenshots before updating the document
    save_chart_screenshots(app)
    generator = DocGenerator(app, app.generated_document_path.text(), app.update_document_path.toolTip())
    generator.generate()
    log_message("Document update completed")