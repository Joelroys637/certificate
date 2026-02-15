import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st
import zipfile
from io import BytesIO


def generate_id_cards_from_excel(template_path, excel_file, y, column_name, text_size, text_color):
    # Load template image
    try:
        template = Image.open(template_path)
    except Exception as e:
        st.error(f"Error loading template: {e}")
        return []

    # Load Font
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    try:
        font = ImageFont.truetype(font_path, size=text_size)  # ✅ Dynamic Font Size
    except Exception as e:
        st.error(f"Error loading font: {e}")
        return []

    # Read Excel File
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return []

    # Validate Column
    if column_name not in df.columns:
        st.error(f"Excel file must contain this column: {column_name}")
        return []

    id_cards = []

    for i, student in df.iterrows():
        card = template.copy()
        draw = ImageDraw.Draw(card)

        text = str(student[column_name])

        # Center text horizontally
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        center_x = (card.width - text_width) // 2

        try:
            text_y = int(y)
            draw.text((center_x, text_y), text, fill=text_color, font=font)  # ✅ Color Applied
            id_cards.append(card)
        except:
            st.error("Please enter only numbers for Y axis!")

    return id_cards


# ---------------- STREAMLIT APP ----------------

def main():
    st.title("Certificate Generator App")

    column_name = st.text_input("Enter Excel Column Name (Ex: name)")

    template_file = st.file_uploader("Upload Certificate Template", type=["png", "jpg", "jpeg"])
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    st.write("Find your certificate coordinates here:")
    st.write("https://imageonline.io/find-coordinates-of-image/")

    y = st.text_input("Enter the Y axis position (only number):")

    # ✅ NEW FEATURES
    text_size = st.slider("Select Text Size", min_value=20, max_value=150, value=60)
    text_color = st.color_picker("Pick Text Color", "#000000")

    if template_file and excel_file and column_name and y:
        template_path = "uploaded_template.png"

        with open(template_path, "wb") as f:
            f.write(template_file.getvalue())

        id_cards = generate_id_cards_from_excel(
            template_path,
            excel_file,
            y,
            column_name,
            text_size,      # ✅ Added
            text_color      # ✅ Added
        )

        if id_cards:
            st.success(f"Generated {len(id_cards)} Certificates Successfully!")

            zip_buffer = BytesIO()

            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for i, card in enumerate(id_cards):
                    img_buffer = BytesIO()
                    card.save(img_buffer, format="PNG")
                    img_buffer.seek(0)

                    zf.writestr(f"certificate_{i + 1}.png", img_buffer.read())

                    st.image(card, caption=f"Certificate {i + 1}", use_column_width=True)

            st.download_button(
                label="Download All Certificates as ZIP",
                data=zip_buffer.getvalue(),
                file_name="certificates.zip",
                mime="application/zip"
            )


main()
