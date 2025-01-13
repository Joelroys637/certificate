import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st
import os
import zipfile
from io import BytesIO

def generate_id_cards_from_excel(template_path, excel_file,y):
    """
    Generate ID cards with student details from an Excel file on a given template.

    Args:
        template_path (str): Path to the ID card template image.
        excel_file (UploadedFile): Uploaded Excel file containing 'name' column.

    Returns:
        List of generated ID card images (in-memory).
    """
    # Load the template image
    try:
        template = Image.open(template_path)
    except Exception as e:
        st.error(f"Error loading template: {e}")
        return []

    # Set up font (provide a valid font file path)
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"  # Replace with a valid font file path on your system
    try:
        font = ImageFont.truetype(font_path, size=100)
    except Exception as e:
        st.error(f"Error loading font: {e}")
        return []

    # Read the Excel file
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return []

    # Check for required columns
    if 'name' not in df.columns:
        st.error("Excel file must contain a 'name' column.")
        return []

    # Generate ID cards
    id_cards = []
    for i, student in df.iterrows():
        # Copy the template to draw on
        card = template.copy()
        draw = ImageDraw.Draw(card)

        # Calculate text width and center it
        text = f"{student['name']}"
        text_bbox = draw.textbbox((0, 0), text, font=font)  # Get text bounding box
        text_width = text_bbox[2] - text_bbox[0]  # Width of the text
        center_x = (card.width - text_width) // 2  # Horizontally center the text

        try:
            text_y = int(y)  # Vertical position (adjust as needed)
            draw.text((center_x, text_y), text, fill="black", font=font)
            id_cards.append(card)
        except:
            pass
        # Add text to the template
        

    return id_cards

# Streamlit app
def main():
    st.title("Certificate Generator")

    # Upload template image
    template_file = st.file_uploader("Upload certificate Template (Image)", type=["png", "jpg", "jpeg"])
    excel_file = st.file_uploader("Upload Excel File (with 'name' column)", type=["xlsx"])


    st.write("Find in your certificate Coordinates usning this link:'https://pixspy.com/'")
    st.image("example.png")
    st.write("Findout the y Coordinate in your certificate and Enter the following box you don't Worry about x axis because the name is fixed in the y axis center")
    y=st.text_input("Enter the Y axis in certificate template image NOTE(only Enter number): ")
    
    if template_file and excel_file:
        # Save the template locally to process
        template_path = "uploaded_template.png"
        with open(template_path, "wb") as f:
            f.write(template_file.getvalue())

        # Generate ID cards
        id_cards = generate_id_cards_from_excel(template_path, excel_file,y)

        if id_cards:
            st.success(f"Generated {len(id_cards)} Certificate!")
            
            # Show generated images and prepare for download
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for i, card in enumerate(id_cards):
                    img_buffer = BytesIO()
                    card.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    zf.writestr(f"id_card_{i + 1}.png", img_buffer.read())

                    # Display each ID card
                    st.image(card, caption=f"ID Card {i + 1}", use_column_width=True)

            # Add download button
            st.download_button(
                label="Download All ID Cards as ZIP",
                data=zip_buffer.getvalue(),
                file_name="id_cards.zip",
                mime="application/zip"
            )
main()
