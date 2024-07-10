import sys
from docx2txt import process
from presidio_analyzer import AnalyzerEngine
from presidio_anonymizer import AnonymizerEngine, OperatorConfig
from docx import Document


def should_redact(entity, text):
    # Whitelist of words that should not be redacted
    whitelist = ["Azure", "JupyterLab", "Git", "Matplotlib", "Gunicorn", "Tableau"]

    # Check if the entity text is in the whitelist
    if (entity.entity_type == "LOCATION" or entity.entity_type == "PERSON") and text[
        entity.start : entity.end
    ] in whitelist:
        return False

    # Add more conditions here if needed
    return True


if len(sys.argv) != 2:
    print("Usage: python script.py <input_file>")
    sys.exit(1)

input_file = sys.argv[1]

# Extract text from Word document
text = process(input_file)

# Initialize Presidio Analyzer
analyzer = AnalyzerEngine()

# Analyze the text for PII entities
analyzer_results = analyzer.analyze(text=text, language="en")

# Filter out DATE_TIME and IN_PAN entities
filtered_results = [
    result
    for result in analyzer_results
    if result.entity_type not in ["DATE_TIME", "IN_PAN"] and should_redact(result, text)
]

# Create a custom operator configuration
operator_config = {
    "PERSON": OperatorConfig("replace", {"new_value": "Harry Potter"}),
    "PHONE_NUMBER": OperatorConfig("replace", {"new_value": "437-543-3244"}),
    "EMAIL_ADDRESS": OperatorConfig("replace", {"new_value": "harry@gmail.com"}),
}

# Initialize Presidio Anonymizer
anonymizer = AnonymizerEngine()

# Redact the text
redacted_text = anonymizer.anonymize(
    text=text, analyzer_results=filtered_results, operators=operator_config
)

# Construct the output file name
output_file = input_file.rsplit(".", 1)[0] + "-redacted.docx"

# Write the redacted text to a new Word document
try:
    new_doc = Document()
    new_doc.add_paragraph(redacted_text.text)
    new_doc.save(output_file)
    print(f"Redacted document saved as {output_file}")
except Exception as e:
    print(f"Error saving redacted document: {e}")
    sys.exit(1)
