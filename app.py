import streamlit as st
import pandas as pd
import numpy as np
import openpyxl # Do I need this? Remove to test at some point
import math
from io import BytesIO

# Set a fixed random seed for reproducibility.
np.random.seed(42)

def assign_poster_boards(posters, days=2):
    """
    Shuffle posters and assign each one a day and board number.
    The simple logic here assigns the first half to Day 1 and the rest to Day 2.
    """
    posters = posters.sample(frac=1).reset_index(drop=True)
    n = len(posters)
    # Create a Day column: first n//days for Day 1 and the remainder for Day 2.
    posters['Day'] = ['Day 1'] * (n // days) + ['Day 2'] * (n - n // days)
    # Assign board numbers within each day.
    posters['Board_Number'] = posters.groupby('Day').cumcount() + 1

    # Add Session column (AM/PM)
    posters.insert(
        posters.columns.get_loc('Day') + 1,
        'Session',
        posters['Board_Number'].apply(lambda x: 'AM' if x % 2 == 1 else 'PM')
    )

    # For robust sorting (especially if more than 2 days), extract the numeric part of the Day.
    posters['Day_Num'] = posters['Day'].str.extract(r'(\d+)').astype(int)
    # Map Session to sort order: AM first, then PM
    posters['Session_order'] = posters['Session'].map({'AM': 0, 'PM': 1})

    # Sort by Day number, then by Session and finally by Board_Number
    posters = posters.sort_values(
        by=['Day_Num', 'Session_order', 'Board_Number']
    ).drop(['Day_Num', 'Session_order'], axis =1) # Drop the Day_Num and Session_order columns that were only used for sorting

    return posters

def assign_judges(posters, judges, reviews_per_poster):
    """
    Assign judges to posters in a load-balanced way.
    For each poster, select eligible judges (those not from the same lab, if possible)
    and pick the ones with the fewest assignments so far.

    Returns two DataFrames:
    - poster_assignments_df: each poster's assignment with separate judge columns.
    - judge_assignments_df: for each judge, a list of assigned posters with details.
    """
    # Initialize judge load and assignments dictionaries
    judge_load = {judge: 0 for judge in judges['Name']}
    judge_assignments = {judge: [] for judge in judges['Name']}
    assignment_list = []
    
    for idx, poster in posters.iterrows():
        poster_lab = poster['Lab']
        # Try to exclude judges from the same lab.
        eligible_judges = judges[judges['Lab'] != poster_lab]['Name'].tolist()
        if len(eligible_judges) < reviews_per_poster:
            # If not enough eligible judges, use all judges.
            eligible_judges = judges['Name'].tolist()

        # Sort eligible judges by current load (lowest first)
        sorted_judges = sorted(eligible_judges, key=lambda j: judge_load[j])
        # Select the first judges up to the number of "reviews_per_poster"
        selected_judges = sorted_judges[:reviews_per_poster]

        # Check if we have enough judges to assign.
        if len(selected_judges) < reviews_per_poster:
            error_message = (
                f"Error: Not enough judges available for poster '{poster['Poster_Title']}' "
                f"on {poster['Day']} at Board {poster['Board_Number']}. "
                f"Required {reviews_per_poster} judges, but only {len(selected_judges)} were found. "
                "Please add more judges or adjust the eligibility criteria."
            )
            raise ValueError(error_message)

        # Update judge load and record assignment for each selected judge
        for judge in selected_judges:
            judge_load[judge] += 1
            judge_assignments[judge].append({
                'Poster_Title': poster['Poster_Title'],
                'Day': poster['Day'],
                'Session': poster['Session'],
                'Board_Number': poster['Board_Number']
            })

        # Build the assigned dictionary for the current poster
        assignment = {
            'Day': poster['Day'],
            'Session': poster['Session'],
            'Board_Number': poster['Board_Number'],
            'FirstName': poster['FirstName'],
            'LastName': poster['LastName']
        }

        # Insert judge columns immediately after 'LastName'
        for i in range(1, reviews_per_poster + 1):
            assignment[f'Judge_{i}'] = selected_judges[i-1]

        # Append the remaining fields.
        assignment.update({
            'Lab': poster['Lab'],
            'Poster_Title': poster['Poster_Title']
        })

        assignment_list.append(assignment) 
    
    # Make it a pandas DF for easier Excel formatting later
    poster_assignments_df = pd.DataFrame(assignment_list)

    # Build the judge assignments DataFrame.
    judge_assignments_list = []
    for judge, assignments in judge_assignments.items():
        assignment_strs = [
            f"{a['Day']} (Board {a['Board_Number']})" for a in assignments
        ]
        judge_assignments_list.append({
            'Judge': judge,
            'Assigned_Posters': ",".join(assignment_strs)
        })
    judge_assignments_df = pd.DataFrame(judge_assignments_list)

    return poster_assignments_df, judge_assignments_df

def generate_excel(poster_assignments_df, judge_assigments_df, presenters_df, judges_df):
    """
    Generate an Excel workbook (in memory) with four sheets:
     Sheet 1: "Poster Assignments" (the new output)
     Sheet 2: "Judge Review Assignments" (the mapping of the judges to assigned posters)
     Sheet 3: "Original Presenter" (the original presenters dataframe)
     Sheet 4: "Original Judges" (the original judges dataframe)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        poster_assignments_df.to_excel(writer, sheet_name="Poster Assignments", index=False)
        judge_assigments_df.to_excel(writer, sheet_name="Judge Review Assignments", index=False)
        presenters_df.to_excel(writer, sheet_name="Original Presenter", index=False)
        judges_df.to_excel(writer, sheet_name="Original Judges", index=False)

        # auto-adjust column widths for each sheet.
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                # Get the column letter (e.g. A, B C)
                column = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                # Add a little extra space for readability
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column].width = adjusted_width
    processed_data = output.getvalue()
    return processed_data

# ---------------------
# Streamlit UI components
# ---------------------

st.title("Poster & Judge Assignment Tool")
st.markdown("""
This tool assigns poster boards and judge reviews for your poster presentation event.
Upload the Excel files with poster presenter details and judge details, set the number of reviews per poster,
and click **Generate Assignments** to produce an Excel file with:
 - Poster assignments (load balanced across judges)
 - A mapping of judges to the posters they will review.
 - Original presenter data
 - Original judge data
""")

# File uploader widgets for the two Excel files.
excel_file = st.file_uploader("Upload Poster Presenters and Judges Excel File", type=["xlsx"])

# Configurable parameter for number of reviews per poster.
reviews_per_poster = st.number_input("Number of Reviews per Poster", min_value=1, value=2, step=1)

# Select which sheet is what
# Error handling to present error messages here as a problem with excel handling
# Implemented specifically because excel handling wrecked me for half a day aaaahhhhhhhhh
# Or it might have been a misspelled comma in my spec file IDK anymore
# Hello anybody reading this! Sorry for that rant. signed. Tomoya
if excel_file is not None:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names

        poster_sheet = st.selectbox("Select the sheet for Poster Presenters", sheet_names)
        judge_sheet = st.selectbox("Select the sheet for Judges", sheet_names)

        presenters_df = pd.read_excel(excel_file, sheet_name=poster_sheet)
        judges_df = pd.read_excel(excel_file, sheet_name=judge_sheet)
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")

if st.button("Generate Assignments"):
    # Send error if user doesn't actually upload anything before clicking go.
    if excel_file is None:
        st.error("Please upload an Excel file with one page for presenters and one page for judges.")
    else:
        # Check that required columns are present.
        required_poster_cols = {"FirstName", "LastName", "Lab", "Poster_Title"}
        required_judge_cols = {"Name", "Lab"}
        if not required_poster_cols.issubset(presenters_df.columns):
            st.error(f"Presenter sheet must contain these columns: {required_poster_cols}")
        elif not required_judge_cols.issubset(judges_df.columns):
            st.error(f"Judge sheet must contain these columns: {required_judge_cols}")
        else:
            try:
                # Step 1: Assign poster boards
                presenters_df = assign_poster_boards(presenters_df, days=2)
                # Step 2: Assign judges using load balancing
                poster_assignments_df, judge_assignments_df = assign_judges(presenters_df, judges_df, reviews_per_poster)
                # Step 3: Generate Excel workbook with four sheets
                excel_data = generate_excel(poster_assignments_df, judge_assignments_df, presenters_df, judges_df)
                
                st.success("Assignments generated successfully!")
                st.download_button(
                    label="Download Assignment Excel",
                    data=excel_data,
                    file_name="poster_judge_assignments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Physical Boards needed calculation
                max_board = presenters_df['Board_Number'].max()
                physical_boards = math.ceil(max_board / 2)
                st.write(f"Maximum Physical Boards Needed: {physical_boards}")

                # Total Posters/Judges numbers displayed
                num_posters = len(presenters_df)
                num_judges = len(judges_df)
                st.write(f"Total Posters: {num_posters}")
                st.write(f"Total Judges: {num_judges}")
                
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")