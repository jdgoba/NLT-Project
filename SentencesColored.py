import spacy
from docx import Document
from docx.shared import RGBColor

def colorize_word(run, pos):
# Generate a unique color for each POS tag
    color = hash(pos) % 0x1000000  # Use hash value as RGB color
    run.font.color.rgb = RGBColor(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF)


def process_text(text, doc):
    # Load the SpaCy model
    nlp = spacy.load('en_core_web_sm')

    # Process the text
    docx_paragraph = doc.add_paragraph()
    spacy_doc = nlp(text)

    for sent in spacy_doc.sents:

        # Analyze the sentence and colorize words
        for token in sent:
            run = docx_paragraph.add_run(str(token.text) + " ")
            colorize_word(run, token.pos_)

        # Add a line break between sentences
        docx_paragraph.add_run('\n')

if __name__ == "__main__":
    # Get user input
    input_text = "This is a text. He eats an apple."

    # Create a Word document
    doc = Document()

    # Process the text and add sentences to the Word document
    process_text(input_text, doc)

    # Save the Word document
    doc.save('output.docx')

    print("Word document saved successfully.")
