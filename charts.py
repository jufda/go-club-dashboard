import altair as alt
import pandas as pd


#######################
# Win and loss chart
#######################
def make_win_loss_chart(input_df, input_player):
    """Create a win/loss chart for a player."""
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

#######################
# Expected vs actual wins
#######################
def make_expected_vs_actual_chart(filtered_df, selected_player):
    """Create a chart comparing expected vs actual wins."""
    if selected_player == "ALL PLAYERS":
        expected_wins = filtered_df['Vahvemman voiton todennäköisyys'].sum()
        actual_wins = filtered_df[filtered_df['Voittaja'] == filtered_df['Pelaaja vahvempi']].shape[0]
    else:
        expected_wins = filtered_df['Selected Player Win Probability'].sum()
        actual_wins = filtered_df[filtered_df['Voittaja'] == selected_player].shape[0]

    win_data = pd.DataFrame({
        'Type': ['Expected', 'Actual    '],
        'Count': [expected_wins, actual_wins]
    })

    chart = alt.Chart(win_data).mark_bar().encode(
        x=alt.X('Type', title=''),
        y=alt.Y('Count', title=''),
        color=alt.Color('Type', legend=None),
        tooltip=alt.Tooltip('Count', format='.2f')
    ).properties(width=150, height=300)

    return chart

#######################
# Activity timeline
#######################
def make_performance_chart(input_df, input_player):
    """Create a activity timeline chart for a player."""
    if input_player == "ALL PLAYERS":
        input_df = input_df.melt(
            id_vars=['Päivämäärä'],
            value_vars=['Pelaaja vahvempi', 'Pelaaja heikompi'],
            var_name='Role',
            value_name='Player'
        )
        grouped_df = input_df.groupby(['Päivämäärä', 'Player']).size().reset_index(name='Game Count')
    else:
        input_df = input_df[(input_df['Pelaaja vahvempi'] == input_player) | (input_df['Pelaaja heikompi'] == input_player)]
        input_df['Opponent'] = input_df.apply(
            lambda row: row['Pelaaja heikompi'] if row['Pelaaja vahvempi'] == input_player else row['Pelaaja vahvempi'],
            axis=1
        )
        grouped_df = input_df.groupby(['Päivämäärä', 'Opponent']).size().reset_index(name='Game Count')

    chart = alt.Chart(grouped_df).mark_bar().encode(
        x=alt.X('Päivämäärä:T', title='Date'),
        y=alt.Y('Game Count:Q', title='Number of players'),
        color=alt.Color('Player:N' if input_player == "ALL PLAYERS" else 'Opponent:N', title='Player'),
        tooltip=['Päivämäärä:T', 'Player:N' if input_player == "ALL PLAYERS" else 'Opponent:N', 'Game Count:Q']
    ).properties(
        width=800,
        height=300,
        title=f"{input_player}'s club activity timeline"
    ).interactive()

    return chart

#######################
# Rating timeline
#######################
def make_rating_timeline_chart(input_df, input_player, selected_opponent="NONE"):
    """Create a rating timeline chart for player(s)."""
    if input_player == "ALL PLAYERS":
        return None
    
    rating_data_player1 = []
    rating_data_player2 = []

    # Get all games for player 1 (selected player)
    stronger_games_p1 = input_df[input_df['Pelaaja vahvempi'] == input_player]
    for _, row in stronger_games_p1.iterrows():
        rating_data_player1.append({
            'Date': row['Päivämäärä'],
            'Rating': row['Rating vahv'],
            'Player': input_player
        })

    weaker_games_p1 = input_df[input_df['Pelaaja heikompi'] == input_player]
    for _, row in weaker_games_p1.iterrows():
        rating_data_player1.append({
            'Date': row['Päivämäärä'],
            'Rating': row['Rating heik'],
            'Player': input_player
        })

    if selected_opponent != "NONE":
        # Get all games for player 2 (opponent), not just games against selected player
        stronger_games_p2 = input_df[input_df['Pelaaja vahvempi'] == selected_opponent]
        for _, row in stronger_games_p2.iterrows():
            rating_data_player2.append({
                'Date': row['Päivämäärä'],
                'Rating': row['Rating vahv'],
                'Player': selected_opponent
            })

        weaker_games_p2 = input_df[input_df['Pelaaja heikompi'] == selected_opponent]
        for _, row in weaker_games_p2.iterrows():
            rating_data_player2.append({
                'Date': row['Päivämäärä'],
                'Rating': row['Rating heik'],
                'Player': selected_opponent
            })
    
    if not rating_data_player1:
        return None
        
    rating_df_player1 = pd.DataFrame(rating_data_player1)
    rating_df_player1 = rating_df_player1.sort_values('Date')
    
    all_rating_data = [rating_df_player1]
    if rating_data_player2:
        rating_df_player2 = pd.DataFrame(rating_data_player2)
        rating_df_player2 = rating_df_player2.sort_values('Date')
        all_rating_data.append(rating_df_player2)

    combined_rating_df = pd.concat(all_rating_data)
    
    min_rating = combined_rating_df['Rating'].min()
    max_rating = combined_rating_df['Rating'].max()
    rating_padding = (max_rating - min_rating) * 0.1 if (max_rating - min_rating) > 0 else 10
    y_min = min_rating - rating_padding
    y_max = max_rating + rating_padding
    
    rating_ticks = list(range(int(y_min // 100) * 100, int(y_max // 100 + 1) * 100, 100))
    if not rating_ticks and combined_rating_df['Rating'].nunique() == 1:
        rating_ticks = [int(combined_rating_df['Rating'].iloc[0])]
    elif not rating_ticks:
        rating_ticks = [2000, 2100]

    rank_df = pd.DataFrame({'value': rating_ticks})
    rank_df['rank'] = rank_df['value'].apply(lambda x: 
        f"{int((x - 2000) // 100)}d" if x >= 2100 else f"{int((2100 - x) // 100)}k"
    )
    
    base = alt.Chart(combined_rating_df).encode(
        x=alt.X('Date:T', title='Date')
    )
    line = base.mark_line().encode(
        y=alt.Y('Rating:Q',
                scale=alt.Scale(domain=[y_min, y_max], nice=False),
                axis=alt.Axis(
                    title='Rating',
                    values=rating_ticks,
                    grid=True
                )),
        color='Player:N'
    )
    
    points = base.mark_point(size=50).encode(
        y=alt.Y('Rating:Q',
                scale=alt.Scale(domain=[y_min, y_max], nice=False)),
        tooltip=['Date:T', alt.Tooltip('Rating:Q', format='.0f'), 'Player:N'],
        color='Player:N'
    )

    rating_labels = alt.Chart(pd.DataFrame({'value': rating_ticks})).mark_text(
        align='right',
        baseline='middle',
        dx=-60
    ).encode(
        y=alt.Y('value:Q', scale=alt.Scale(domain=[y_min, y_max], nice=False)),
        text=alt.Text('value:Q', format='.0f'),
        color=alt.Color('rank:N', scale=alt.Scale(scheme='category10'), legend=None)
    )
    
    rank_axis_chart = alt.Chart(rank_df).mark_text(
        align='right',
        baseline='middle',
        dx=60,
        fontWeight='bold'
    ).encode(
        y=alt.Y('value:Q',
                scale=alt.Scale(domain=[y_min, y_max], nice=False),
                axis=alt.Axis(
                    orient='right',
                    title='',
                    values=rating_ticks,
                    grid=False
                )),
        text='rank:N',
        color=alt.Color('rank:N', scale=alt.Scale(scheme='category10'), legend=None)
    )
    
    chart_title = f"{input_player}'s Rating Timeline"
    if selected_opponent != "NONE":
        chart_title += f" vs {selected_opponent}"

    chart = (line + points + rating_labels + rank_axis_chart).properties(
        width=800,
        height=300,
        title=chart_title
    ).interactive()
    
    return chart