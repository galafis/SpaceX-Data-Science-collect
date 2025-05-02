import sys
sys.path.append("/opt/.manus/.sandbox-runtime")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN # Ensure PP_ALIGN is imported
import importlib.util

# --- Import appendix chart definitions ---
try:
    spec = importlib.util.spec_from_file_location("slide_content_module", "/home/ubuntu/slide_content_en.py")
    slide_content_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(slide_content_module)
    appendix_charts = slide_content_module.appendix_charts
    print("Successfully imported appendix chart definitions.")
except Exception as e:
    print(f"Error importing appendix chart definitions: {e}")
    appendix_charts = [] # Define empty list to avoid errors

# --- Paths ---
presentation_path = "/home/ubuntu/spacex_presentation_watermarked_en.pptx" # Input with watermark
output_path = "/home/ubuntu/spacex_presentation_appendix_en.pptx" # Final output

# --- Helper function to add caption ---
def add_caption(slide, text, left, top, width, height):
    try:
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = text
        p = tf.paragraphs[0]
        p.font.size = Pt(10)
        p.alignment = PP_ALIGN.CENTER # Use imported PP_ALIGN
    except Exception as e:
        print(f"Error adding caption 	'{text}	': {e}")

# --- Main script ---
try:
    prs = Presentation(presentation_path)
    print(f"Loaded presentation: {presentation_path}")

    # --- Find Appendix Title Slide Index ---
    apendice_title_idx = -1
    for i, slide in enumerate(prs.slides):
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            if "appendix" in slide.shapes.title.text.lower():
                apendice_title_idx = i
                print(f"Found Appendix title slide at index {i}")
                break

    if apendice_title_idx == -1:
        print("Warning: Appendix title slide not found. Adding charts at the end.")
        insert_idx = len(prs.slides)
    else:
        insert_idx = apendice_title_idx + 1

    # --- Add Appendix Chart Slides ---
    # Use a consistent layout, e.g., Title and Content (index 1)
    try:
        content_layout = prs.slide_layouts[1]
    except IndexError:
        print("Error: Title and Content layout (index 1) not found. Using layout 0.")
        content_layout = prs.slide_layouts[0] # Fallback

    # Add one slide per chart definition
    for chart_info in appendix_charts:
        try:
            # Add slide at the end first, then handle reordering if possible/needed
            # For simplicity and robustness, we add at the end.
            slide = prs.slides.add_slide(content_layout)
            chart_title = chart_info.get(	'title	', 	'Untitled Chart	')
            print(f"Added new slide for appendix chart: {chart_title}")

            # Add Title
            title_shape = slide.shapes.title
            if title_shape:
                title_shape.text = chart_info.get("title", "Appendix Chart")
            else:
                print("Warning: Title placeholder not found on appendix chart slide.")

            # Add Image
            img_path = chart_info.get("image_path")
            if img_path:
                # Position the image (adjust as needed)
                left = Inches(1.0)
                top = Inches(1.5)
                height = Inches(4.5) # Adjust size
                try:
                    pic = slide.shapes.add_picture(img_path, left, top, height=height)
                    print(f"Added image: {img_path}")

                    # Add Caption below the image
                    caption_text = chart_info.get("caption", "")
                    if caption_text:
                        caption_top = top + height + Inches(0.1)
                        caption_width = pic.width
                        caption_height = Inches(0.5)
                        add_caption(slide, caption_text, left, caption_top, caption_width, caption_height)
                        print(f"Added caption: {caption_text}")

                except FileNotFoundError:
                    print(f"Error: Image file not found at {img_path}")
                except Exception as e_img:
                    print(f"Error adding image {img_path}: {e_img}")
            else:
                print("Warning: No image path specified for this appendix chart.")

        except Exception as e_slide:
            chart_title_error = chart_info.get(	'title	', 	'Unknown Chart	') # Get title safely
            print(f"Error adding slide for chart 	'{chart_title_error}	': {e_slide}")

    # --- Save the final presentation ---
    prs.save(output_path)
    print(f"Presentation with appendix (English) saved to: {output_path}")

except Exception as e:
    print(f"An error occurred during appendix addition: {e}")


