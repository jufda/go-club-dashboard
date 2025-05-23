#######################
# Import libraries
#######################
import altair as alt
from datetime import datetime, timedelta
import os
import pandas as pd
from pandas import to_datetime
import requests
import streamlit as st
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

#######################
# Page configuration
#######################
st.set_page_config(
    page_title="Go Club Games Dashboard",
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

# 1st season
try:
    df1 = pd.read_excel(
        io='data/goseason1.xlsx',
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:E,G,P,Q,W",
        skiprows=3
    )
except Exception as e:
    st.warning(f"Failed to load data from local source for Season 1. Using remote file with a typoed year. Error: {e}")
    df1 = pd.read_excel(
        io="https://docs.google.com/spreadsheets/d/1NImVfaJ3z7K_hrlvFsC-HmmtXIj7x2ZVcIZthPf2IYs/export?format=xlsx",
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:E,G,P,Q,W",
        skiprows=3
    )

# 2nd season
try:
    df2 = pd.read_excel(
        io='data/goseason2.xlsx',
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:F,O,P,V",
        skiprows=3
    )
except Exception as e:
    st.warning(f"Failed to load data from local source for Season 2. Using remote file. Error: {e}")
    df2 = pd.read_excel(
        io="https://docs.google.com/spreadsheets/d/18qYVirc-_ni1I6myhICqA3hMgxP9pfvuBNJ0oLbljIE/export?format=xlsx",
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:F,O,P,V",
        skiprows=3
    )

# current season
local_file = "data/latest-season.xlsx"
online_file_url = "https://docs.google.com/spreadsheets/d/1ktWkll-JubPHH3CSbcqO22WaJci2yAaQSK7-UKG8do8/export?format=xlsx"

if os.path.exists(local_file):
    try:
        last_modified_time = datetime.fromtimestamp(os.path.getmtime(local_file))
        current_time = datetime.now()

        if current_time - last_modified_time > timedelta(days=3):
            print("Local file is older than 3 days. Downloading the latest version...")
            download_file(online_file_url, local_file)
        # else:
            # print("Local file seems fresh, using it.")
    except Exception as e:
        print(f"Error checking online file: {e}. Using the local file.")
else:
    print("Local file not found. Downloading the latest version...")
    download_file(online_file_url, local_file)

try:
    df3 = pd.read_excel(
        io="data/latest-season.xlsx",
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:F,O,P,V",
        skiprows=3
    )
except Exception as e:
    st.warning(f"Failed to load data from online source for Season 3. Using local backup file. Error: {e}")
    df3 = pd.read_excel(
        io="data/goseason3.xlsx",
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:F,O,P,V",
        skiprows=3
    )

#######################
# Clean data
#######################
df1 = df1.set_axis(['Pelaaja vahvempi', 'Pelaaja heikompi', 'Tasoituskivet', 'Voittaja', 'Päivämäärä', 'Rating vahv', 'Rating heik', 'Vahvemman voiton todennäköisyys'], axis=1)
df2 = df2.set_axis(['Pelaaja vahvempi', 'Pelaaja heikompi', 'Tasoituskivet', 'Voittaja', 'Päivämäärä', 'Rating vahv', 'Rating heik', 'Vahvemman voiton todennäköisyys'], axis=1)
df3 = df3.set_axis(['Pelaaja vahvempi', 'Pelaaja heikompi', 'Tasoituskivet', 'Voittaja', 'Päivämäärä', 'Rating vahv', 'Rating heik', 'Vahvemman voiton todennäköisyys'], axis=1)

df1.dropna(axis=0, inplace=True)
df2.dropna(axis=0, inplace=True)
df3.dropna(axis=0, inplace=True)

df1['Päivämäärä'] = to_datetime(df1['Päivämäärä'], format='mixed', dayfirst=True)
df2['Päivämäärä'] = to_datetime(df2['Päivämäärä'], format='mixed', dayfirst=True)
df3['Päivämäärä'] = to_datetime(df3['Päivämäärä'], format='mixed', dayfirst=True)

df = pd.concat([df1, df2, df3], ignore_index=True)
df.sort_values(by=['Päivämäärä', 'Pelaaja vahvempi', 'Pelaaja heikompi'], ascending=True, inplace=True)


#######################
# Sidebar
#######################
with st.sidebar:
    st.title('Go Club Games Dashboard')

    # Player selection
    players = df['Pelaaja vahvempi'].unique().tolist() + df['Pelaaja heikompi'].unique().tolist()
    players = sorted(set(players))
    players.insert(0, "ALL PLAYERS")  # Add "ALL PLAYERS" at the top of the list
    selected_player = st.selectbox('Select a player', players, index=0)  # Default to "ALL PLAYERS"

    # Date range selection
    min_date = df['Päivämäärä'].min()
    max_date = df['Päivämäärä'].max()
    selected_date_range = st.date_input('Select a date range', [min_date, max_date])

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


#######################
# Plots
#######################

# Win/loss ratio
def make_win_loss_chart(input_df, input_player):
    if input_player == "ALL PLAYERS":
        wins = input_df[input_df['Voittaja'] == input_df['Pelaaja vahvempi']].shape[0]
        losses = input_df.shape[0] - wins
    else:
        wins = input_df[input_df['Voittaja'] == input_player].shape[0]
        losses = input_df[(input_df['Pelaaja vahvempi'] == input_player) |
                         (input_df['Pelaaja heikompi'] == input_player)].shape[0] - wins
    data = pd.DataFrame({
        'Result': ['Wins', 'Losses'],
        'Count': [wins, losses]
    })
    chart = alt.Chart(data).mark_bar().encode(
        x=alt.X('Result', title=''),
        y=alt.Y('Count', title=''),
        color = alt.Color('Result', legend=None)
    ).properties(width=150, height=300)
    return chart


# Expected vs actual wins
def make_expected_vs_actual_chart():
    if selected_player == "ALL PLAYERS":
        expected_wins = filtered_df['Vahvemman voiton todennäköisyys'].sum()
        actual_wins = filtered_df[filtered_df['Voittaja'] == filtered_df['Pelaaja vahvempi']].shape[0]
    else:
        expected_wins = filtered_df['Selected Player Win Probability'].sum()
        actual_wins = filtered_df[filtered_df['Voittaja'] == selected_player].shape[0]

    # Create a DataFrame for the chart
    win_data = pd.DataFrame({
        'Type': ['Expected', 'Actual    '],
        'Count': [expected_wins, actual_wins]
    })

    # Create the bar chart
    chart = alt.Chart(win_data).mark_bar().encode(
        x=alt.X('Type', title=''),
        y=alt.Y('Count', title=''),
        color=alt.Color('Type', legend=None),
        tooltip=alt.Tooltip('Count', format='.2f')
    ).properties(width=150, height=300)

    return chart


# Player performance timeline with colourful opponents
def make_performance_chart(input_df, input_player):
    if input_player == "ALL PLAYERS":
        # Include both players in each game
        input_df = input_df.melt(
            id_vars=['Päivämäärä'],
            value_vars=['Pelaaja vahvempi', 'Pelaaja heikompi'],
            var_name='Role',
            value_name='Player'
        )
        grouped_df = input_df.groupby(['Päivämäärä', 'Player']).size().reset_index(name='Game Count')
    else:
        # Filter data for the selected player
        input_df = input_df[(input_df['Pelaaja vahvempi'] == input_player) | (input_df['Pelaaja heikompi'] == input_player)]

        # Add a column for the opponent
        input_df['Opponent'] = input_df.apply(
            lambda row: row['Pelaaja heikompi'] if row['Pelaaja vahvempi'] == input_player else row['Pelaaja vahvempi'],
            axis=1
        )

        # Group by date and opponent, then count games
        grouped_df = input_df.groupby(['Päivämäärä', 'Opponent']).size().reset_index(name='Game Count')

    # Create a bar chart
    chart = alt.Chart(grouped_df).mark_bar().encode(
        x=alt.X('Päivämäärä:T', title='Date'),
        y=alt.Y('Game Count:Q', title='Number of players'),
        color=alt.Color('Player:N' if input_player == "ALL PLAYERS" else 'Opponent:N', title='Player'),
        tooltip=['Päivämäärä:T', 'Player:N' if input_player == "ALL PLAYERS" else 'Opponent:N', 'Game Count:Q']
    ).properties(
        width=800,
        height=300,
        title=f"{input_player}'s club activity timeline"
    ).interactive()  # Enable zoom and pan

    return chart


# Rating timeline chart
def make_rating_timeline_chart(input_df, input_player):
    if input_player == "ALL PLAYERS":
        return None
    
    # Create a list to store rating data
    rating_data = []
    
    # Get games where player was stronger
    stronger_games = input_df[input_df['Pelaaja vahvempi'] == input_player]
    for _, row in stronger_games.iterrows():
        rating_data.append({
            'Date': row['Päivämäärä'],
            'Rating': row['Rating vahv']
        })
    
    # Get games where player was weaker
    weaker_games = input_df[input_df['Pelaaja heikompi'] == input_player]
    for _, row in weaker_games.iterrows():
        rating_data.append({
            'Date': row['Päivämäärä'],
            'Rating': row['Rating heik']
        })
    
    # Convert to DataFrame and sort by date
    if not rating_data:
        return None
        
    rating_df = pd.DataFrame(rating_data)
    rating_df = rating_df.sort_values('Date')
    
    # Calculate min and max ratings for the y-axis domain
    min_rating = rating_df['Rating'].min()
    max_rating = rating_df['Rating'].max()
    # Add some padding to the range
    rating_padding = (max_rating - min_rating) * 0.1
    y_min = min_rating - rating_padding
    y_max = max_rating + rating_padding
    
    # Create rating ticks for both rating and rank scales
    rating_ticks = list(range(int(y_min // 100) * 100, int(y_max // 100 + 1) * 100, 100))
    
    # Create a DataFrame for the rank axis with pre-calculated ranks
    rank_df = pd.DataFrame({'value': rating_ticks})
    rank_df['rank'] = rank_df['value'].apply(lambda x: 
        f"{int((x - 2000) // 100)}d" if x >= 2100 else f"{int((2100 - x) // 100)}k"
    )
    
    # Create the base chart with rating on left y-axis
    base = alt.Chart(rating_df).encode(
        x=alt.X('Date:T', title='Date')
    )
    
    # Create the line and points
    line = base.mark_line().encode(
        y=alt.Y('Rating:Q',
                scale=alt.Scale(domain=[y_min, y_max]),
                axis=alt.Axis(
                    title='Rating',
                    # tickCount=len(rating_ticks),
                    values=rating_ticks,
                    grid=True
                ))
    )
    
    points = base.mark_point(size=50).encode(
        y=alt.Y('Rating:Q',
                scale=alt.Scale(domain=[y_min, y_max])),
        tooltip=[alt.Tooltip('Date:T', title='Date'),
             alt.Tooltip('Rating:Q', title='Rating', format=".1f")]
    )

    # Create the rank axis on the right using the pre-calculated ranks
    rank_axis = alt.Chart(rank_df).mark_text(
        align='right',
        baseline='middle',
        dx=15,  # Offset from the right edge
        fontWeight='bold'
    ).encode(
        y=alt.Y('value:Q',
                scale=alt.Scale(domain=[y_min, y_max]),
                axis=alt.Axis(
                    orient='right',
                    title='Rank',
                    values=rating_ticks,
                    grid=False
                )),
        text='rank:N',
        color=alt.Color('rank:N', scale=alt.Scale(scheme='category10'), legend=None)
    )
    
    # Create a layered chart
    chart = alt.layer(line, points, rank_axis).resolve_scale(
        y='independent'
    ).properties(
        width=800,
        height=300,
        title=f"{input_player}'s rating timeline"
    ).interactive()
    
    return chart


#######################
# Dashboard Top
#######################
col = st.columns((8, 1.5, 1.5), gap='medium')

with col[0]:
    # st.markdown("""#### Player's club activity timeline""")
    performance_chart = make_performance_chart(filtered_df, selected_player)
    st.altair_chart(performance_chart, use_container_width=True)
    
    # Add rating timeline chart if a specific player is selected
    if selected_player != "ALL PLAYERS":
        # st.markdown("""""")
        rating_chart = make_rating_timeline_chart(filtered_df, selected_player)
        if rating_chart:
            st.altair_chart(rating_chart, use_container_width=True)

with col[1]:
    st.markdown('#### Games')
    win_loss_chart = make_win_loss_chart(filtered_df, selected_player)
    st.altair_chart(win_loss_chart, use_container_width=True)

with col[2]:
    st.markdown('#### Wins')
    expected_vs_actual_chart = make_expected_vs_actual_chart()
    st.altair_chart(expected_vs_actual_chart, use_container_width=True)

#######################
# Dashboard Main
#######################
with st.container():
    st.markdown('#### Game details')
    sorted_df = filtered_df.sort_values(by="Päivämäärä", ascending=False)
    st.dataframe(
        sorted_df,
        hide_index=False,
        column_order=("Päivämäärä", "Weekday", "Pelaaja vahvempi", "Vahvemman voiton todennäköisyys", "Pelaaja heikompi",
                      "Rating vahv", "Tasoituskivet", "Rating heik", "Voittaja"),
        column_config={
            "Pelaaja vahvempi": "Player (Stronger)",  # Rename column
            "Vahvemman voiton todennäköisyys": st.column_config.ProgressColumn(
                "Stronger Win Probability", format="%.2f"  # Format as percentage
            ),
            "Pelaaja heikompi": "Player (Weaker)",  # Rename column
            "Tasoituskivet": st.column_config.NumberColumn(
                "Handicap Stones", format="%d"  # Format as integer
            ),
            "Voittaja": "Winner",  # Rename column
            "Päivämäärä": st.column_config.DateColumn(
                "Date", format="YYYY-MM-DD"  # Format as date
            ),
            "Rating vahv": st.column_config.NumberColumn(
                "Stronger Rating", format="%.0f"  # Format as float with 2 decimals
            ),
            "Rating heik": st.column_config.NumberColumn(
                "Weaker Rating", format="%.0f"  # Format as float with 2 decimals
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
- **Select opponent**: to allow comparison between two players.
- **Head2head**: win%, games played against each other. y=games x=player 
- **Estimations with confidence intervals**: Calculate the error margins based on the amount of games so far.
- **Handicap analysis**: How players perform with handicap stones.
- **Translation**: Offer both Finnish and English
- **Fix bugs**: Error while selecting date range
- **All players'** timeline: colour based on the amount of games
- **Accessibility testing**: Colour-blindness, contrast, alt-texts, keyboard navigation.
- **Visual look**: Include go art, board, stones and cups to create a fitting theme.
""")