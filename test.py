import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO

# Set a fixed random seed (you might let the user choose this later)
np.random.seed(42)

def assign_poster_boards(posters, days=2):
    """
    Shuffle posters and assign each one a day and board number.
    """
    posters = posters.sample(frac=1).reset_index(drop=True)
    n = len(posters)
    posters['Day'] = ['Day 1'] * (n // days) + ['Day 2'] * (n - n // days)
    posters['Board_Number'] = posters.groupby('Day').cumcount() + 1
    return posters

def assign_judges_load_balanced(posters, judges, reviews_per_poster):
    """
    Assign judges to posters in a load-balanced way.
    For each poster, select eligible judges (those not from the same lab, if possible)
    and pick the ones with the fewest assignments so far.
    Returns two DataFrames:
      - poster_assignments_df: each poster's assignment with a comma-separated list of judges.
      - judge_assignments_df: for each judge, a list of assigned posters with details.
    """
    # Initialize judge load and assignments dictionaries
    judge_load = {judge: 0 for judge in judges['Name']}
    judge_assignments = {judge: [] for judge in judges['Name']}
    assignment_list = []
   
    for idx, poster in posters.iterrows():
        poster_lab = poster['Lab']
        # Try to exclude judges from the same lab
        eligible_judges = judges[judges['Lab'] != poster_lab]['Name'].tolist()
        if len(eligible_judges) < reviews_per_poster:
            eligible_judges = judges['Name'].tolist()
       
        # Sort eligible judges by current load (lowest first)
        sorted_judges = sorted(eligible_judges, key=lambda j: judge_load[j])
        # Select the first 'reviews_per_poster' judges
        selected_judges = sorted_judges[:reviews_per_poster]
       
        # Update judge load and record assignment for each judge
        for judge in selected_judges:
            judge_load[judge] += 1
            judge_assignments[judge].append({
                'Poster_Title': poster['Poster_Title'],
                'Presenter': poster['Name'],
                'Lab': poster['Lab'],
                'Day': poster['Day'],
                'Board_Number': poster['Board_Number']
            })
       
        assignment_list.append({
            'Poster_Title': poster['Poster_Title'],
            'Presenter': poster['Name'],
            'Lab': poster['Lab'],
            'Day': poster['Day'],
            'Board_Number': poster['Board_Number'],
            'Assigned_Judges': ", ".join(selected_judges)
        })
   
    poster_assignments_df = pd.DataFrame(assignment_list)
   
    # Build the judge assignments DataFrame:
    judge_assignments_list = []
    for judge, assignments in judge_assignments.items():
        # Create a summary string for each assignment
        assignment_strs = [
            f"{a['Poster_Title']} (Board {a['Board_Number']}, {a['Day']})"
            for a in assignments
        ]
        judge_assignments_list.append({
            'Judge': judge,
            'Assigned_Posters': ", ".join(assignment_strs)
        })
    judge_assignments_df = pd.DataFrame(judge_assignments_list)
   
    return poster_assignments_df, judge_assignments_df

def generate_excel(poster_assignments_df, judge_assignments_df, presenters_df, judges_df):
    """
    Generate an Excel workbook (in memory) with four sheets:
      Sheet 1: "Poster Assignments" (the new output)
      Sheet 2: "Original Presenters" (the original presenters dataframe)
      Sheet 3: "Original Judges" (the original judges dataframe)
      Sheet 4: "Judge Review Assignments" (the mapping of judges to assigned posters)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        poster_assignments_df.to_excel(writer, sheet_name="Poster Assignments", index=False)
        presenters_df.to_excel(writer, sheet_name="Original Presenters", index=False)
        judges_df.to_excel(writer, sheet_name="Original Judges", index=False)
        judge_assignments_df.to_excel(writer, sheet_name="Judge Review Assignments", index=False)
    processed_data = output.getvalue()
    return processed_data

# ---------------------
# Streamlit UI Components
# ---------------------

st.title("Poster & Judge Assignment Tool with Load Balancing")
st.markdown("""
This tool assigns poster boards and judge reviews for your poster presentation event.
Upload the Excel file with presenter and judge details, set the number of reviews per poster,
and click **Generate Assignments** to produce an Excel file with:
  - Poster assignments (load balanced across judges)
  - Original presenter data
  - Original judge data
  - A mapping of judges to the posters they will review.
""")

# File uploader widget for the Excel file.
excel_file = st.file_uploader("Upload Excel File (Presenters and Judges)", type=["xlsx"])

# Configurable parameter for the number of reviews per poster.
reviews_per_poster = st.number_input("Number of Reviews per Poster", min_value=1, value=2, step=1)

if excel_file is not None:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names

        poster_sheet = st.selectbox("Select the sheet for Presenter Details", sheet_names)
        judge_sheet = st.selectbox("Select the sheet for Judge Details", sheet_names)

        presenters_df = pd.read_excel(excel_file, sheet_name=poster_sheet)
        judges_df = pd.read_excel(excel_file, sheet_name=judge_sheet)
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")

if st.button("Generate Assignments"):
    if excel_file is None:
        st.error("Please upload an Excel file with the required sheets.")
    else:
        # Check for required columns
        required_presenter_cols = {"Name", "Lab", "Poster_Title"}
        required_judge_cols = {"Name", "Lab"}
        if not required_presenter_cols.issubset(presenters_df.columns):
            st.error(f"Presenter sheet must contain columns: {required_presenter_cols}")
        elif not required_judge_cols.issubset(judges_df.columns):
            st.error(f"Judge sheet must contain columns: {required_judge_cols}")
        else:
            try:
                # Step 1: Assign poster boards
                presenters_df = assign_poster_boards(presenters_df, days=2)
                # Step 2: Assign judges using load balancing
                poster_assignments_df, judge_assignments_df = assign_judges_load_balanced(presenters_df, judges_df, reviews_per_poster)
                # Step 3: Generate Excel workbook with four sheets
                excel_data = generate_excel(poster_assignments_df, judge_assignments_df, presenters_df, judges_df)
               
                st.success("Assignments generated successfully!")
                st.download_button(
                    label="Download Assignment Excel",
                    data=excel_data,
                    file_name="poster_judge_assignments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")