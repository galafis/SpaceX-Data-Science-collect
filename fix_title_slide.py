import sys
sys.path.append("/opt/.manus/.sandbox-runtime")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Paths
input_pptx = "/home/ubuntu/spacex_presentation_appendix_en.pptx" # Input from previous step
output_pptx = "/home/ubuntu/spacex_presentation_final_title_fixed_en.pptx" # Final PPTX
author_name = "Gabriel Demetrios Lafis"
date_str = "May 2, 2025"

try:
    prs = Presentation(input_pptx)
    print(f"Loaded presentation: {input_pptx}")

    if len(prs.slides) > 0:
        slide1 = prs.slides[0]
        print("Accessing first slide to add author/date text box.")

        # Define position and size for the new text box
        # These values might need adjustment based on the template layout
        left = Inches(1.0)  # Adjust as needed
        top = Inches(5.5)   # Adjust as needed (below the main title)
        width = Inches(6.0) # Adjust as needed
        height = Inches(0.5) # Adjust as needed

        # Add the text box
        txBox = slide1.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = f"{author_name}\n{date_str}"

        # Optional: Format the text (adjust font, size, color, alignment)
        for paragraph in tf.paragraphs:
            paragraph.alignment = PP_ALIGN.LEFT # Or CENTER/RIGHT
            for run in paragraph.runs:
                run.font.name = 'Calibri' # Match template font if known
                run.font.size = Pt(14)    # Adjust size
                # Set color (e.g., white if background is dark)
                # run.font.color.rgb = RGBColor(255, 255, 255)

        print(f"Added text box with author and date to the first slide.")

    else:
        print("Warning: Presentation has no slides.")

    # Save the modified presentation
    prs.save(output_pptx)
    print(f"Presentation with fixed title slide saved to: {output_pptx}")

except Exception as e:
    print(f"An error occurred while fixing the title slide: {e}")

