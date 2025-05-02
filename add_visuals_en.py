import sys
sys.path.append("/opt/.manus/.sandbox-runtime")
from pptx import Presentation
from pptx.util import Inches, Pt
import importlib.util

# --- Import slide content definitions (for reference, not strictly needed here) ---
try:
    spec = importlib.util.spec_from_file_location("slide_content_module", "/home/ubuntu/slide_content_en.py")
    slide_content_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(slide_content_module)
    slide_content_en = slide_content_module.slide_content_en
    appendix_charts = slide_content_module.appendix_charts # Get appendix chart info
    print("Successfully imported English slide content definitions.")
except Exception as e:
    print(f"Warning: Could not import slide content definitions: {e}")
    slide_content_en = [] # Define empty list to avoid errors later
    appendix_charts = []

# --- Paths ---
presentation_path = "/home/ubuntu/spacex_presentation_draft_en.pptx"
output_path = "/home/ubuntu/spacex_presentation_visuals_en.pptx"

# Image paths (ensure these exist)
img_config_chart = "/home/ubuntu/rocket_configurations_chart_en.png"
img_launch_sites = "/home/ubuntu/launch_sites_chart_en.png"
img_outcomes_chart = "/home/ubuntu/launch_outcomes_chart_en.png"
img_landings_chart = "/home/ubuntu/booster_landings_chart_en.png"

# --- Helper function to find slide by title (approximate match) ---
def find_slide_by_title(prs, title_keyword):
    for i, slide in enumerate(prs.slides):
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            if title_keyword.lower() in slide.shapes.title.text.lower():
                print(f"Found slide 	'{slide.shapes.title.text}	' (index {i}) matching keyword 	'{title_keyword}	'")
                return slide
    print(f"Warning: Slide with title containing 	'{title_keyword}	' not found.")
    return None

# --- Main script ---
try:
    prs = Presentation(presentation_path)
    print(f"Loaded presentation: {presentation_path}")

    # --- Add images to specific slides ---

    # 1. Rocket Configurations Chart -> Falcon 9 Evolution slide
    evo_slide = find_slide_by_title(prs, "Evolution")
    if evo_slide:
        try:
            # Position near the bottom or side
            left = Inches(5.5)
            top = Inches(4.0)
            height = Inches(2.5) # Adjust size as needed
            pic = evo_slide.shapes.add_picture(img_config_chart, left, top, height=height)
            print(f"Added {img_config_chart} to Evolution slide.")
        except Exception as e:
            print(f"Error adding {img_config_chart}: {e}")

    # 2. Launch Outcomes & Sites Charts -> Launch Statistics slide
    stats_slide = find_slide_by_title(prs, "Statistics")
    if stats_slide:
        try:
            # Position side-by-side or stacked
            left1 = Inches(5.0)
            top1 = Inches(2.0)
            height1 = Inches(2.0)
            pic1 = stats_slide.shapes.add_picture(img_outcomes_chart, left1, top1, height=height1)
            print(f"Added {img_outcomes_chart} to Statistics slide.")

            left2 = Inches(5.0)
            top2 = Inches(4.5) 
            height2 = Inches(2.0)
            pic2 = stats_slide.shapes.add_picture(img_launch_sites, left2, top2, height=height2)
            print(f"Added {img_launch_sites} to Statistics slide.")
        except Exception as e:
            print(f"Error adding charts to Statistics slide: {e}")

    # 3. Booster Landings Chart -> Reusability slide
    reuse_slide = find_slide_by_title(prs, "Reusability")
    if reuse_slide:
        try:
            # Position near the bottom or side
            left = Inches(5.5)
            top = Inches(4.0)
            height = Inches(2.5) # Adjust size as needed
            pic = reuse_slide.shapes.add_picture(img_landings_chart, left, top, height=height)
            print(f"Added {img_landings_chart} to Reusability slide.")
        except Exception as e:
            print(f"Error adding {img_landings_chart}: {e}")

    # Note: Appendix charts will be handled in a separate step/script if needed,
    # based on the appendix_charts definition imported earlier.

    # --- Save the presentation with visuals ---
    prs.save(output_path)
    print(f"Presentation with visuals saved to: {output_path}")

except Exception as e:
    print(f"An error occurred during visual addition: {e}")


