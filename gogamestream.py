#######################
# Import libraries
#######################
import altair as alt
from datetime import datetime, timedelta
import os
import pandas as pd
import requests
import streamlit as st
import warnings
from charts import (
    make_performance_chart,
    make_rating_timeline_chart,
    make_win_loss_chart,
    make_expected_vs_actual_chart,
    make_head_to_head_win_loss_chart,
    make_head_to_head_expected_vs_actual_chart,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

#######################
# Page configuration
#######################
st.set_page_config(
    page_title="Go club games dashboard",
    page_icon="🌑",
    layout="wide",
    initial_sidebar_state="expanded")

alt.theme.enable("dark")

#######################
# Load data
#######################
def download_file(url, local_path):
    """Download a file from a URL and save it locally."""
    response = requests.get(url)
    response.raise_for_status()  # Raise an error for bad responses
    with open(local_path, 'wb') as file:
        file.write(response.content)

@st.cache_data(ttl=timedelta(days=1))
def load_season_data(season_number, local_path, online_url, usecols, skiprows=3):
    """
    Load and clean data for a specific season.
    
    Args:
        season_number (int): The season number (1, 2, or 3)
        local_path (str): Path to the local Excel file
        online_url (str): URL to the online Excel file
        usecols (str): Columns to use from the Excel file
        skiprows (int): Number of rows to skip
        
    Returns:
        pd.DataFrame: Cleaned dataframe for the season
    """
    try:
        # Try loading from local file first
        df = pd.read_excel(
            io=local_path,
            engine="openpyxl",
            sheet_name="Pelitulokset",
            usecols=usecols,
            skiprows=skiprows
        )
    except Exception as e:
        st.warning(f"Failed to load data from local source for Season {season_number}. Using remote file. Error: {e}")
        try:
            df = pd.read_excel(
                io=online_url,
                engine="openpyxl",
                sheet_name="Pelitulokset",
                usecols=usecols,
                skiprows=skiprows
            )
        except Exception as e:
            st.error(f"Failed to load data for Season {season_number}. Error: {e}")
            return pd.DataFrame()  # Return empty DataFrame if both attempts fail
    
    # Clean the data
    df = df.set_axis(['Pelaaja vahvempi', 'Pelaaja heikompi', 'Tasoituskivet', 'Voittaja', 
                     'Päivämäärä', 'Rating vahv', 'Rating heik', 'Vahvemman voiton todennäköisyys'], axis=1)
    df.dropna(axis=0, inplace=True)
    df['Päivämäärä'] = pd.to_datetime(df['Päivämäärä'], format='mixed', dayfirst=True)
    
    return df

# Load data for each season
df1 = load_season_data(
    season_number=1,
    local_path='data/goseason1.xlsx',
    online_url="https://docs.google.com/spreadsheets/d/1NImVfaJ3z7K_hrlvFsC-HmmtXIj7x2ZVcIZthPf2IYs/export?format=xlsx",
    usecols="B:E,G,P,Q,W"
)

df2 = load_season_data(
    season_number=2,
    local_path='data/goseason2.xlsx',
    online_url="https://docs.google.com/spreadsheets/d/18qYVirc-_ni1I6myhICqA3hMgxP9pfvuBNJ0oLbljIE/export?format=xlsx",
    usecols="B:F,O,P,V"
)

# Handle current season with special logic for file freshness
local_file = "data/latest-season.xlsx"
online_file_url = "https://docs.google.com/spreadsheets/d/1ktWkll-JubPHH3CSbcqO22WaJci2yAaQSK7-UKG8do8/export?format=xlsx"

if os.path.exists(local_file):
    try:
        last_modified_time = datetime.fromtimestamp(os.path.getmtime(local_file))
        current_time = datetime.now()

        if current_time - last_modified_time > timedelta(days=3):
            print("Local file is older than 3 days. Downloading the latest version...")
            download_file(online_file_url, local_file)
    except Exception as e:
        print(f"Error checking online file: {e}. Using the local file.")
else:
    print("Local file not found. Downloading the latest version...")
    download_file(online_file_url, local_file)

df3 = load_season_data(
    season_number=3,
    local_path="data/latest-season.xlsx",
    online_url="data/goseason3.xlsx",  # Fallback to local backup
    usecols="B:F,O,P,V"
)

# Combine all seasons
df = pd.concat([df1, df2, df3], ignore_index=True)
df.sort_values(by=['Päivämäärä', 'Pelaaja vahvempi', 'Pelaaja heikompi'], ascending=True, inplace=True)

#######################
# Sidebar
#######################
with st.sidebar:
    st.title('Go club games dashboard')

    # Player selection
    players = df['Pelaaja vahvempi'].unique().tolist() + df['Pelaaja heikompi'].unique().tolist()
    players = sorted(set(players))
    players.insert(0, "ALL PLAYERS")  # Add "ALL PLAYERS" at the top of the list
    selected_player = st.selectbox('Select player', players, index=0)  # Default to "ALL PLAYERS"

    # Opponent selection - only show players who have played against the selected player
    if selected_player != "ALL PLAYERS":
        # Get all games where the selected player participated
        player_games = df[
            (df['Pelaaja vahvempi'] == selected_player) |
            (df['Pelaaja heikompi'] == selected_player)
        ]
        # Get unique opponents from these games
        opponents = set()
        for _, row in player_games.iterrows():
            if row['Pelaaja vahvempi'] == selected_player:
                opponents.add(row['Pelaaja heikompi'])
            else:
                opponents.add(row['Pelaaja vahvempi'])
        available_opponents = sorted(list(opponents))
    else:
        available_opponents = []
    
    available_opponents.insert(0, "NONE")  # Option to not select an opponent
    selected_opponent = st.selectbox('Select opponent', available_opponents, index=0)

    # Date range selection
    min_date = df['Päivämäärä'].min()
    max_date = df['Päivämäärä'].max()
    selected_date_range = st.date_input('Select date range', [min_date, max_date])

    # Filter data based on selections
    filtered_df = df[(df['Päivämäärä'] >= pd.to_datetime(selected_date_range[0])) &
                     (df['Päivämäärä'] <= pd.to_datetime(selected_date_range[1]))]

    if selected_player != "ALL PLAYERS":
        filtered_df = filtered_df[(filtered_df['Pelaaja vahvempi'] == selected_player) |
                                  (filtered_df['Pelaaja heikompi'] == selected_player)]

    # Add a column for the day of the week
    filtered_df['Weekday'] = df['Päivämäärä'].dt.day_name()

    # Calculate win probability for the selected player if one is selected
    if selected_player:
        filtered_df['Selected Player Win Probability'] = filtered_df.apply(
            lambda row: row['Vahvemman voiton todennäköisyys'] if row['Pelaaja vahvempi'] == selected_player
            else 1 - row['Vahvemman voiton todennäköisyys'], axis=1
        )
    else:
        filtered_df['Selected Player Win Probability'] = None

    # Create a separate dataframe for game details that can be filtered by opponent
    game_details_df = filtered_df.copy()
    if selected_opponent != "NONE":
        game_details_df = game_details_df[
            ((game_details_df['Pelaaja vahvempi'] == selected_player) & (game_details_df['Pelaaja heikompi'] == selected_opponent)) |
            ((game_details_df['Pelaaja vahvempi'] == selected_opponent) & (game_details_df['Pelaaja heikompi'] == selected_player))
        ]

#######################
# Dashboard Top
#######################
col = st.columns((8, 1.5, 1.5), gap='medium')

with col[0]:
    performance_chart = make_performance_chart(filtered_df, selected_player)
    st.altair_chart(performance_chart, use_container_width=True)
    
    # Add rating timeline chart if a specific player is selected
    if selected_player != "ALL PLAYERS":
        rating_chart = make_rating_timeline_chart(filtered_df, selected_player, selected_opponent)
        if rating_chart:
            st.altair_chart(rating_chart, use_container_width=True)

with col[1]:
    st.markdown('#### All games')
    win_loss_chart = make_win_loss_chart(filtered_df, selected_player)
    st.altair_chart(win_loss_chart, use_container_width=True)
    
    # Add head-to-head win/loss chart if an opponent is selected
    if selected_player != "ALL PLAYERS" and selected_opponent != "NONE":
        h2h_win_loss_chart = make_head_to_head_win_loss_chart(filtered_df, selected_player, selected_opponent)
        st.altair_chart(h2h_win_loss_chart, use_container_width=True)

with col[2]:
    st.markdown('#### All wins')
    expected_vs_actual_chart = make_expected_vs_actual_chart(filtered_df, selected_player)
    st.altair_chart(expected_vs_actual_chart, use_container_width=True)
    
    # Add head-to-head expected vs actual wins chart if an opponent is selected
    if selected_player != "ALL PLAYERS" and selected_opponent != "NONE":
        h2h_expected_vs_actual_chart = make_head_to_head_expected_vs_actual_chart(filtered_df, selected_player, selected_opponent)
        st.altair_chart(h2h_expected_vs_actual_chart, use_container_width=True)

#######################
# Dashboard Main
#######################
with st.container():
    st.markdown('#### Game details')
    sorted_df = game_details_df.sort_values(by="Päivämäärä", ascending=False)
    st.dataframe(
        sorted_df,
        hide_index=False,
        column_order=("Päivämäärä", "Weekday", "Rating vahv", "Pelaaja vahvempi", "Vahvemman voiton todennäköisyys",
                      "Rating heik","Pelaaja heikompi", "Tasoituskivet", "Voittaja"),
        column_config={
            "Pelaaja vahvempi": "Player (stronger)",
            "Vahvemman voiton todennäköisyys": st.column_config.ProgressColumn(
                "Stronger player's win %", format="%.2f"
            ),
            "Pelaaja heikompi": "Player (weaker)",
            "Tasoituskivet": st.column_config.NumberColumn(
                "Handicap stones", format="%d"
            ),
            "Voittaja": "Winner",  # Rename column
            "Päivämäärä": st.column_config.DateColumn(
                "Date", format="YYYY-MM-DD"
            ),
            "Rating vahv": st.column_config.NumberColumn(
                "Rating (s.)", format="%.0f"
            ),
            "Rating heik": st.column_config.NumberColumn(
                "Rating (w.)", format="%.0f"
            ),

        }

    )

#######################
# Bottom info
#######################
    with st.expander('About the stats for the selected player', expanded=False):
        st.write('''
            - **Player's club games timeline**: Displays the activity in club games over time. ALL PLAYERS view show the number of games played.
            - **Games**: Shows the number of wins and losses. ALL PLAYERS view shows stats based on the stronger-by-rating player of each game.
            - **Wins**: Shows the number of expected wins based or players' ratings and handicap stones compared to actual wins. ALL PLAYERS view shows the stats based on the stronger-by-rating player of each game.
            - **Game details**: Lists all player's recorded club games and provides statistics.
            ''')


    with st.expander('Update history', expanded=False):
        st.write("""
#### Updates 26.5.2025:
1. **Opponent selection**: Shows only the games between players.
2. **Performance optimization**: Caches the data.
3. **Face-to-face comparison**: Compare the rating timelines of two players.
4. **Face-to-face wins**: Bar charts for head-to-head comparison.
        
#### Updates 23.5.2025:
1. **Rating graph**: shows player's rating over time
2. **Download optimization**: checks if the local file is older than 3 days before downloading a new version.
3. **Included all players to timeline**: not just half of them.
4. **Ordered game details**: newest first.
5. **Decimal formatting**: for expected wins and rating.

#### Updates 15.5.2025:
1. **Timeline update**: shows opponents
2. **Added expected wins**: and comparison to actual wins
3. **Rearranged dame details**: win probability visualized and column order improved 
4. **Added statistics for ALL PLAYERS**: Win/loss and expected/actual wins based on the higher-rating-player of each game.

#### Updates 14.5.2025:
1. **Added third season games**
2. **Included more data**: player's ratings, expected win %s
3. **Updated timeline**: Game colour by winner, barchart based on played game dates

#### Prototype 17.2.2025:
1. **Data Loading from online spreadsheets**: The code now loads and cleans the Go club games data from the online Excel files.
2. **Sidebar**: The sidebar allows users to select a player and a date range to filter the data.
3. **Timeline**: A chart showing games played by the selected player over time.
4. **Win/Loss Ratio**: A bar chart showing the number of wins and losses for the selected player.
5. **Recent games**: Displays the 10 most recent games involving the selected player.

#### Next steps in the further development:
- **Head2head**: Win%, games played against each other. y=games x=player 
- **Estimations with confidence intervals**: Calculate the error margins based on the amount of games so far.
- **Handicap analysis**: How players perform with handicap stones.
- **Translation**: Offer both Finnish and English
- **Fix bugs**: Error while selecting date range, nicer loading
- **Player selection**: From the game details table.
- **All players'** timeline: Player colour based on the amount of games
- **Accessibility testing**: Colour-blindness, contrast, alt-texts, keyboard navigation.
- **Visual look**: Include go art, board, stones and cups to create a fitting theme.
- **Opponent rating graph**: Currently only shows the "selected" vs "opponent" in the graph for the opponent.
""")