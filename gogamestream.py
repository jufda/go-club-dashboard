#######################
# Import libraries
#######################
import streamlit as st
import altair as alt
import pandas as pd
from pandas import to_datetime
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

try:
    df3 = pd.read_excel(
        io='https://docs.google.com/spreadsheets/d/1ktWkll-JubPHH3CSbcqO22WaJci2yAaQSK7-UKG8do8/export?format=xlsx',
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols="B:F,O,P,V",
        skiprows=3
    )
except Exception as e:
    st.warning(f"Failed to load data from online source for Season 3. Using local file. Error: {e}")
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
        color=alt.Color('Type', legend=None)
    ).properties(width=150, height=300)

    return chart


# Player performance timeline with colourful opponents
def make_performance_chart(input_df, input_player):
    # If no player is selected, show data for all players by adding the lost player to list of opponents
    if input_player == "ALL PLAYERS":
        input_df['Opponent'] = input_df.apply(
            lambda row: row['Pelaaja heikompi'] if row['Pelaaja vahvempi'] == row['Voittaja'] else row[
                'Pelaaja vahvempi'],
            axis=1
        )
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
        y=alt.Y('Game Count:Q', title='Number of Games'),
        color=alt.Color('Opponent:N', title='Opponent'),
        tooltip=['Päivämäärä:T', 'Opponent:N', 'Game Count:Q']
    ).properties(
        width=800,
        height=300
    ).interactive()  # Enable zoom and pan

    return chart


#######################
# Dashboard Top
#######################
col = st.columns((8, 1.5, 1.5), gap='medium')

with col[0]:
    st.markdown("""#### Player's club activity timeline""")
    performance_chart = make_performance_chart(filtered_df, selected_player)
    st.altair_chart(performance_chart, use_container_width=True)

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
    st.dataframe(
        filtered_df,
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
1. **More informative visualizations**: handicap analysis, actual wins vs expected wins
2. **Translation**: Offer both Finnish and English
4. **Interactive elements**: Like hover tooltips, click interactions, etc.
5. **Fix bugs**: Error while selecting date rang
6. **Better layout**: Improve the layout and design of the dashboard for better user experience.
7. **Visualize rating**: Add a visualization of player ratings over time.
8. **Deploy**: Create github.io page for the dashboard or share through streamlit
9. **Accessibility testing**: Colour-blindness, contrast, alt-texts, keyboard navigation.
10. **Head2head win% estimations with confidence intervals**: Calculate the error margins based on the amount of games so far.
11. **Visual look**: Include go art, board, stones and cups to create a fitting theme.
12. **Updating to web page**: extending 
""")