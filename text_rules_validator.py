import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from us2uk_QC import load_us_to_uk_dict, detect_us_words

# Load US to UK dictionary
us_uk_dict = load_us_to_uk_dict()

# Pattern to detect contractions with smart (’) and straight (') apostrophes
CONTRACTION_PATTERN = re.compile(
    r"\b(?:I['’]m|don['’]t|can['’]t|won['’]t|shouldn['’]t|could['’]ve|you['’]d|it['’]s|he['’]d|she['’]d|they['’]re|we['’]re|didn['’]t|doesn['’]t|isn['’]t|wasn['’]t|aren['’]t|weren['’]t|['’]re|['’]d|['’]ll|['’]ve|['’]m|['’]s|n['’]t)\b",
    re.IGNORECASE
)

def has_contraction(text):
    return bool(CONTRACTION_PATTERN.search(text))

def has_us_spelling(text):
    return detect_us_words(text, us_uk_dict)

def has_extra_spaces(text):
    return "  " in text

# Improved function to handle any whitespace or hidden characters before checking for period
def has_ending_period(text):
    text = str(text).strip()  # Strip any leading or trailing spaces
    return text.endswith(".")

def has_mid_sentence_period(text):
    return bool(re.search(r"(?<!^)\.(?!$)", text))

# Function to check for inconsistent use of periods across a slide
def highlight_inconsistent_periods(df_qc):
    highlighted_cells = []

    # Grouping rows by Slide Number
    grouped = df_qc.groupby("Slide Number")

    for slide_num, group in grouped:
        # Check if all sentences in the slide either end with a period or not
        all_with_period = all(has_ending_period(row["Extracted Text"]) for _, row in group.iterrows())
        all_without_period = all(not has_ending_period(row["Extracted Text"]) for _, row in group.iterrows())

        for idx, row in group.iterrows():
            ends_with = has_ending_period(row["Extracted Text"])

            # Debugging: Check for inconsistent periods
            if (all_with_period and not ends_with) or (all_without_period and ends_with):
                print(f"Highlighting Slide {slide_num}, Row {idx}: {row['Extracted Text']}")
                highlighted_cells.append((slide_num, idx, "Ending Period Used"))

    return highlighted_cells

def run_text_rules_validation(df_qc):
    if "Extracted Text" not in df_qc.columns:
        return pd.DataFrame([{"Error": "Missing 'Extracted Text' column in QC sheet"}])

    result = []
    for _, row in df_qc.iterrows():
        text = str(row.get("Extracted Text", "")).strip()

        # Skip completely empty lines
        if not text:
            continue

        contraction_found = has_contraction(text)
        us_words = has_us_spelling(text)

        # Adding "Yes" or "No" in the "Ending Period Used" column
        ending_period_used = "Yes" if has_ending_period(text) else "No"

        result.append({
            "File Name": row.get("File Name", ""),
            "Slide Number": row.get("Slide Number", ""),
            "Shape Name / Table Cell": row.get("Shape Name / Table Cell", ""),
            "Extracted Text": text,
            "Contraction Used": "Yes" if contraction_found else "",
            "US English Used": ", ".join(us_words) if us_words else "",
            "Extra Space": "Yes" if has_extra_spaces(text) else "",
            "Ending Period Used": ending_period_used,
            "Mid Sentence Period Used": "Yes" if has_mid_sentence_period(text) else ""
        })

    # Highlight inconsistency in period usage
    highlighted_cells = highlight_inconsistent_periods(df_qc)

    # Update the DataFrame with highlights (coloring cells or adding an indicator)
    for slide_num, idx, column in highlighted_cells:
        df_qc.at[idx, column] = "Highlighted"  # This can be replaced with actual formatting logic

    print(df_qc)  # Debugging: Check if "Highlighted" is being applied correctly

    return pd.DataFrame(result)

# Function to apply red and green color highlighting ONLY to the "Ending Period Used" column (Column H)
def apply_excel_highlighting(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active

    # Define light red and light green color for highlighting
    light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    # Debugging: Ensure we are correctly iterating through rows and checking values
    print("Applying color to cells in column H based on period usage...")

    # Apply the highlight color ONLY to the "Ending Period Used" column (Column H)
    for row in ws.iter_rows(min_row=2, max_col=ws.max_column):
        ending_period_cell = row[7]  # Column H (index 7) for "Ending Period Used"
        
        # Check the value before applying color
        print(f"Checking cell: {ending_period_cell.value}")  # Debugging

        # Highlight with light red if "No" (no period)
        if ending_period_cell.value == "No":
            ending_period_cell.fill = light_red_fill  # Red for inconsistent period

        # Highlight with light green if "Yes" (period present)
        elif ending_period_cell.value == "Yes":
            ending_period_cell.fill = light_green_fill  # Green for consistent period

    # Save the updated file
    wb.save(excel_path)
    print(f"Coloring applied and file saved to {excel_path}")

# Debug mode
if __name__ == "__main__":
    # Load the Excel file
    df = pd.read_excel("S3_MBA_Financial_Statement_Analysis_and_Business_Valuation_10.4_Applications_and_Limitations_of_Relative_Valuation_Techniques_V1_QC_Report (8).xlsx")
    
    # Run the validation
    df_out = run_text_rules_validation(df)
    
    # Export the output to Excel
    df_out.to_excel("text_rules_output.xlsx", index=False)

    # Apply highlighting to the Excel file (only highlight Column H)
    apply_excel_highlighting("text_rules_output.xlsx")
