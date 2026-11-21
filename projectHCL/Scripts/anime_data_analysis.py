# ============================================================
#  PROJECT NAME : Anime Data Analysis
#  TRAINING     : HCL Python Training
#  CONCEPTS USED: Tuple, List, Dictionary, Functions, Loops, Pandas
#  DATA FILE    : anime_data.xlsx (220 anime) - keep it in the same folder
#
#  NOTE: Keep the Excel file in the SAME folder as this .py file.
#        Install the libraries once:  pip install pandas openpyxl
# ============================================================

import os
import pandas as pd


# ------------------------------------------------------------
# STEP 1: Load the data from the Excel file (spreadsheet)
# ------------------------------------------------------------
# Look for the Excel file in the same folder where this .py file is saved.
# LOOP through all files in the folder and pick the .xlsx file with "anime" in its name
folder = os.path.dirname(os.path.abspath(__file__))
file_name = ""
for f in os.listdir(folder):
    name = f.lower()
    if name.endswith(".xlsx") and "anime" in name and "result" not in name and not name.startswith("~$"):
        file_name = os.path.join(folder, f)
        break

if file_name == "":
    print("ERROR: Anime Excel file (.xlsx) not found!")
    print("Copy the Excel file into this folder:", folder)
    print("Files found in this folder:", os.listdir(folder))
    raise SystemExit

print("Reading data from:", file_name)
df = pd.read_excel(file_name, sheet_name="Anime Dataset")

# DICTIONARY: old long column names -> new short names
short_names = {
    "Anime Title": "Title",
    "Primary Genre": "Genre",
    "Release Year": "Year",
    "Source Material": "Source",
    "User Rating (sample)": "Rating",
    "Popularity Rank (sample)": "Rank",
    "Members (sample)": "Members",
}
df = df.rename(columns=short_names)

# TUPLE: columns we want to show in most tables (tuple cannot be changed)
main_columns = ("Title", "Genre", "Year", "Episodes", "Rating")

# LIST: all the different genres in the data
genre_list = list(df["Genre"].unique())


# ------------------------------------------------------------
# STEP 2: Functions for analysis
# ------------------------------------------------------------
def show_data_info():
    print("\n----- ABOUT THE DATA -----")
    print("Rows and columns :", df.shape)
    print("Column names     :", list(df.columns))
    print("\nFirst 5 rows:")
    print(df[list(main_columns)].head())
    print("\nMissing values in each column:")
    print(df.isnull().sum())


def clean_data():
    # Some anime have no episode count (blank in Excel).
    # We count them and then fill them with 0 for simple analysis.
    print("\n----- DATA CLEANING -----")
    missing = df["Episodes"].isnull().sum()
    print("Anime with missing episode count :", missing)
    df["Episodes"] = df["Episodes"].fillna(0)
    df["Episodes"] = df["Episodes"].astype(int)
    print("Missing episodes filled with 0")


def show_basic_info():
    print("\n----- BASIC INFORMATION -----")
    print("Total anime        :", len(df))
    print("Total genres       :", len(genre_list))
    print("Average rating     :", round(df["Rating"].mean(), 2))
    print("Highest rating     :", df["Rating"].max())
    print("Lowest rating      :", df["Rating"].min())
    print("Oldest release year:", df["Year"].min())
    print("Latest release year:", df["Year"].max())
    print("Total members      :", df["Members"].sum())


def top_rated(n):
    print("\n----- TOP", n, "RATED ANIME -----")
    top = df.sort_values(by="Rating", ascending=False).head(n)
    print(top[["Title", "Genre", "Rating"]])


def lowest_rated(n):
    print("\n----- BOTTOM", n, "RATED ANIME -----")
    low = df.sort_values(by="Rating").head(n)
    print(low[["Title", "Genre", "Rating"]])


def most_popular(n):
    print("\n----- TOP", n, "MOST POPULAR ANIME (by members) -----")
    pop = df.sort_values(by="Members", ascending=False).head(n)
    print(pop[["Title", "Members", "Rating"]])


def longest_anime(n):
    print("\n----- TOP", n, "LONGEST ANIME (by episodes) -----")
    longest = df.sort_values(by="Episodes", ascending=False).head(n)
    print(longest[["Title", "Episodes"]])


def filter_by_genre(genre):
    print("\n----- ANIME IN GENRE:", genre, "-----")
    result = df[df["Genre"] == genre]
    if result.empty:
        print("No anime found for this genre.")
    else:
        print("Total found:", len(result))
        print(result[["Title", "Year", "Rating"]].head(10))


def genre_count():
    # DICTIONARY + LOOP to count anime in each genre
    print("\n----- NUMBER OF ANIME IN EACH GENRE -----")
    counts = {}
    for g in df["Genre"]:
        if g in counts:
            counts[g] += 1
        else:
            counts[g] = 1
    for genre, count in counts.items():
        print(genre, ":", count)
    return counts


def average_rating_by_genre():
    print("\n----- AVERAGE RATING BY GENRE -----")
    avg = df.groupby("Genre")["Rating"].mean().round(2)
    print(avg.sort_values(ascending=False))


def format_analysis():
    print("\n----- ANIME BY FORMAT (TV / Movie / OVA / ONA) -----")
    print(df["Format"].value_counts())


def studio_analysis():
    print("\n----- STUDIO ANALYSIS -----")
    print("Anime made by each studio:")
    print(df["Studio"].value_counts())
    print("\nAverage rating of each studio:")
    avg = df.groupby("Studio")["Rating"].mean().round(2)
    print(avg.sort_values(ascending=False))


def source_analysis():
    print("\n----- AVERAGE RATING BY SOURCE MATERIAL -----")
    avg = df.groupby("Source")["Rating"].mean().round(2)
    print(avg.sort_values(ascending=False))


def anime_after_year(year):
    print("\n----- ANIME RELEASED AFTER", year, "-----")
    result = df[df["Year"] > year]
    print("Total found:", len(result))
    print(result[["Title", "Year", "Rating"]].head(10))


def decade_analysis():
    # LOOP + if/elif/else to find the decade of each anime
    print("\n----- ANIME PER DECADE -----")
    decades = []
    for y in df["Year"]:
        if y < 1990:
            decades.append("1980s")
        elif y < 2000:
            decades.append("1990s")
        elif y < 2010:
            decades.append("2000s")
        elif y < 2020:
            decades.append("2010s")
        else:
            decades.append("2020s")
    df["Decade"] = decades
    print(df["Decade"].value_counts().sort_index())


def rating_category():
    # LOOP + if/elif/else to label every anime by rating
    print("\n----- RATING CATEGORY -----")
    categories = []
    for r in df["Rating"]:
        if r >= 8.5:
            categories.append("Masterpiece")
        elif r >= 7.5:
            categories.append("Excellent")
        elif r >= 6.5:
            categories.append("Good")
        else:
            categories.append("Average")
    df["Category"] = categories
    print(df["Category"].value_counts())


def show_graphs():
    # Draw 4 graphs in one window using matplotlib
    # Install once with:  py -m pip install matplotlib
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("\nmatplotlib is not installed, so graphs cannot be shown.")
        print("Run this in the terminal:  py -m pip install matplotlib")
        return

    print("\n----- DRAWING GRAPHS -----")
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    fig.suptitle("Anime Data Analysis", fontsize=16)

    # Graph 1: Bar chart - number of anime in each genre
    genre_total = df["Genre"].value_counts()
    axes[0, 0].bar(genre_total.index, genre_total.values, color="skyblue")
    axes[0, 0].set_title("Number of Anime in Each Genre")
    axes[0, 0].set_ylabel("Number of Anime")
    plt.setp(axes[0, 0].get_xticklabels(), rotation=45, ha="right")

    # Graph 2: Bar chart - average rating of each genre
    genre_avg = df.groupby("Genre")["Rating"].mean().sort_values(ascending=False)
    axes[0, 1].bar(genre_avg.index, genre_avg.values, color="orange")
    axes[0, 1].set_title("Average Rating by Genre")
    axes[0, 1].set_ylabel("Average Rating")
    axes[0, 1].set_ylim(6, 9)
    plt.setp(axes[0, 1].get_xticklabels(), rotation=45, ha="right")

    # Graph 3: Pie chart - anime by format (TV / Movie / OVA / ONA)
    format_total = df["Format"].value_counts()
    axes[1, 0].pie(format_total.values, labels=format_total.index, autopct="%1.1f%%")
    axes[1, 0].set_title("Anime by Format")

    # Graph 4: Bar chart - anime released in each decade
    decade_total = df["Decade"].value_counts().sort_index()
    axes[1, 1].bar(decade_total.index, decade_total.values, color="lightgreen")
    axes[1, 1].set_title("Anime Released in Each Decade")
    axes[1, 1].set_ylabel("Number of Anime")

    plt.tight_layout()
    plt.savefig(os.path.join(folder, "anime_graphs.png"))
    print("Graphs saved as anime_graphs.png")
    plt.show()


def show_conclusion(counts):
    # Find the biggest genre using a LOOP over the dictionary
    biggest_genre = ""
    biggest_count = 0
    for genre, count in counts.items():
        if count > biggest_count:
            biggest_genre = genre
            biggest_count = count

    best = df.loc[df["Rating"].idxmax()]
    popular = df.loc[df["Members"].idxmax()]

    print("\n----- CONCLUSION -----")
    print("Genre with most anime  :", biggest_genre, "(", biggest_count, "anime )")
    print("Highest rated anime    :", best["Title"], "-", best["Rating"])
    print("Most popular anime     :", popular["Title"], "-", popular["Members"], "members")
    print("Best studio (by rating):", df.groupby("Studio")["Rating"].mean().idxmax())


# ------------------------------------------------------------
# STEP 3: Main program - call all the functions
# ------------------------------------------------------------
print("=" * 50)
print("        ANIME DATA ANALYSIS PROJECT")
print("=" * 50)

show_data_info()
clean_data()
show_basic_info()
top_rated(5)
lowest_rated(5)
most_popular(5)
longest_anime(5)
filter_by_genre("Action")
genre_counts = genre_count()
average_rating_by_genre()
format_analysis()
studio_analysis()
source_analysis()
anime_after_year(2015)
decade_analysis()
rating_category()
show_conclusion(genre_counts)
show_graphs()

# Save the final cleaned data into a new Excel file
df.to_excel(os.path.join(folder, "anime_analysis_result.xlsx"), index=False)
print("\nFinal data saved to anime_analysis_result.xlsx")
print("\n----- END OF PROJECT -----")
