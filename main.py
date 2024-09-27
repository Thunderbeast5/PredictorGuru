import pandas as pd
import webbrowser

# Load the Excel file
file_path = 'Final CET.xlsx'
df = pd.read_excel(file_path)

def clean_seat_type(seat_type):
    """
    Remove the first and last characters from the seat type to get the core category.
    """
    if len(seat_type) > 2:
        return seat_type[1:-1]
    else:
        return seat_type

def search_colleges(df, rank=None, percentile=None, seat_type=None, branch=None, district=None, range_=300):
    """
    Search for colleges based on a given rank or percentile with additional filters.
    If the exact rank isn't available, it searches within a range of +/- 'range_' ranks.

    Parameters:
    - df: DataFrame containing the college data.
    - rank: The rank to search for (optional).
    - percentile: The percentile to search for (optional).
    - seat_type: The category or seat type to filter by (optional).
    - branch: The branch or course to filter by (optional).
    - district: The district or location to filter by (optional).
    - range_: The range around the rank to include in the search (optional, default is 300).

    Returns:
    - Filtered DataFrame containing the colleges within the specified filters and rank/percentile range.
    """
    # Apply seat type filter by focusing on the core category
    if seat_type:
        df['CORE SEAT TYPE'] = df['SEAT TYPE'].apply(clean_seat_type)
        df = df[df['CORE SEAT TYPE'] == seat_type]

    # Apply other filters
    if branch:
        df = df[df['BRANCH NAME'] == branch]
    if district:
        df = df[df['DISTRICT'] == district]

    # Filter based on rank or percentile
    if rank is not None:
        df = df[(df['CUT OFF (RANK)'] >= rank - range_) & (df['CUT OFF (RANK)'] <= rank + range_)]
    elif percentile is not None:
        df = df[(df['CUT OFF (PERCENTILE)'] >= percentile - range_) & (df['CUT OFF (PERCENTILE)'] <= percentile + range_)]

    return df

# Example usage with filters
seat_type_filter = 'NT3'  # The core category (e.g., 'OPEN', 'SC', 'OBC')
branch_filter = 'Computer Engineering'  # Filter by branch (e.g., 'Civil Engineering', 'Mechanical Engineering', etc.)
district_filter = 'Nashik'  # Filter by district (e.g., 'Mumbai', 'Pune', etc.)
rank_to_search = 7000  # Rank to search for

# result_df = search_colleges(df, rank=rank_to_search, seat_type=seat_type_filter, branch=branch_filter, district=district_filter)

result_df = search_colleges(df, rank=rank_to_search, seat_type=seat_type_filter)

# Adding CSS styling for better visualization
html = """
<html>
<head>
<style>
    body {{ font-family: Arial, sans-serif; margin: 40px; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
    th, td {{ text-align: left; padding: 8px; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    th {{ background-color: #4CAF50; color: white; }}
    table, th, td {{ border: 1px solid #ddd; }}
    h2 {{ color: #4CAF50; }}
</style>
</head>
<body>

<h2>Filtered Colleges within the Rank Range: {}</h2>
{}

</body>
</html>
""".format(rank_to_search, result_df.to_html(index=False, escape=False))

# Save the result to an HTML file
with open('output.html', 'w') as f:
    f.write(html)

# Open the HTML file in a web browser
webbrowser.open('output.html')
