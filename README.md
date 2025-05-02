# SpaceX Launch Analysis & Presentation Automation

**Author:** Gabriel Demetrios Lafis
**Date:** May 2, 2025

## Project Overview

This project focuses on analyzing SpaceX's Falcon 9 and Falcon Heavy launch history and automating the creation of a comprehensive presentation based on this analysis. The primary goal was to leverage data science techniques to gather insights from publicly available launch data and present them effectively using a standardized template.

This repository contains the Python scripts developed for data processing and presentation generation, the collected data, intermediate files, visual assets, and the final presentation in both PowerPoint (.pptx) and PDF (.pdf) formats.

## Objectives

*   Collect and structure historical data on SpaceX Falcon 9 and Falcon Heavy launches.
*   Analyze key metrics such as launch frequency, success rates, booster landings, and site usage.
*   Develop Python scripts to automatically populate a provided PowerPoint template with the analyzed data, text summaries, and relevant visuals.
*   Generate charts and tables to visualize launch trends and statistics.
*   Ensure the final presentation adheres to specific formatting requirements, including watermarking and an appendix for supplementary materials.
*   Produce a final, polished presentation suitable for review and assessment.

## Methodology

1.  **Data Collection:** Launch data was primarily gathered from the publicly available information on Wikipedia's "List of Falcon 9 and Falcon Heavy launches" page (as of early May 2025). Key information included launch dates, configurations, sites, outcomes, and booster landing attempts/successes.
2.  **Data Structuring:** The collected information was initially structured in a Markdown file (`spacex_data_en.md`) for reference during script development.
3.  **Content Definition:** Key insights and textual summaries for each presentation slide were defined in a separate Python file (`slide_content_en.py`) to decouple content from the generation logic.
4.  **Presentation Automation:** A series of Python scripts using the `python-pptx` library were developed to automate the presentation creation process:
    *   `fill_presentation_en.py`: Populates the base template (`ds-capstone-template-coursera.pptx`) with titles and body text defined in `slide_content_en.py`.
    *   `add_visuals_en.py`: Inserts relevant charts (saved as PNG files from the data source) into appropriate slides.
    *   `add_watermark_en.py`: Adds a watermark with the author's name to every slide.
    *   `fix_title_slide.py`: Manually adds the author's name and date to the first slide via a text box, overcoming template limitations.
    *   `add_appendix_en.py`: Creates an appendix section and adds supplementary charts with captions.
5.  **Visual Assets:** Charts visualizing rocket configurations, launch sites, outcomes, and booster landings were saved directly from the Wikipedia source page.
6.  **Conversion:** The final PowerPoint presentation was converted to PDF format using `unoconv` for wider accessibility.

## Repository Contents

*   **`/` (Root Directory):**
    *   `README.md`: This file.
    *   `fill_presentation_en.py`: Script to populate the presentation template with text content.
    *   `add_visuals_en.py`: Script to add charts/images to the presentation.
    *   `add_watermark_en.py`: Script to add a watermark to all slides.
    *   `add_appendix_en.py`: Script to add the appendix section with charts.
    *   `fix_title_slide.py`: Script to add author/date text box to the first slide.
    *   `slide_content_en.py`: Defines the text content for each slide.
    *   `spacex_data_en.md`: Markdown file with structured data notes (for reference).
    *   `ds-capstone-template-coursera.pptx`: The original PowerPoint template used.
    *   `spacex_presentation_final_title_fixed_en.pptx`: The final generated presentation in PowerPoint format.
    *   `spacex_presentation_final_en.pdf`: The final generated presentation in PDF format.
    *   `*.png`, `*.jpeg`: Image files (charts, etc.) used in the presentation.

## Key Findings & Results

The analysis and resulting presentation highlight the significant increase in launch cadence achieved by SpaceX, the evolution of the Falcon 9 rocket (particularly the Block 5 variant), the remarkable success rate of launches and booster landings contributing to reusability, and the utilization patterns of different launch sites. The automated workflow demonstrates an efficient way to generate data-driven reports and presentations.

## How to Use

1.  Review the final presentation files (`.pptx` and `.pdf`).
2.  Examine the Python scripts (`.py`) to understand the automation logic.
3.  Refer to the `spacex_data_en.md` file for the structured data notes used.

This project serves as a practical example of applying data analysis and automation techniques to real-world scenarios, specifically within the exciting domain of space exploration.

