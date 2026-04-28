import altair as alt
import pandas as pd


# ──────────────────────────────────────────────
# Shared helpers
# ──────────────────────────────────────────────

WIN_COLOR = "#00FFD0"
LOSS_COLOR = "#FF6900"
ACTUAL_COLOR = "#00FFD0"
EXPECTED_COLOR = "#00BBFF"
SMALL_CHART_W = 150
SMALL_CHART_H = 300
LARGE_CHART_H = 300


def _win_loss_bar(data: pd.DataFrame, title: str = "", width: int = SMALL_CHART_W) -> alt.Chart:
    """Shared helper: render a Wins/Losses bar chart from a {'Result','Count'} DataFrame."""
    return (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X("Result", title="", sort=["Wins", "Losses"]),
            y=alt.Y("Count", title=""),
            color=alt.Color(
                "Result",
                legend=None,
                scale=alt.Scale(domain=["Wins", "Losses"], range=[WIN_COLOR, LOSS_COLOR]),
            ),
            tooltip=[
                alt.Tooltip("Result:N"),
                alt.Tooltip("Count:Q", title="Games"),
            ],
        )
        .properties(width=width, height=SMALL_CHART_H, title=title)
    )


def _expected_vs_actual_bar(data: pd.DataFrame, title: str = "") -> alt.Chart:
    """Shared helper: render an Expected/Actual bar chart from a {'Type','Count'} DataFrame."""
    return (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X("Type", title=""),
            y=alt.Y("Count", title=""),
            color=alt.Color(
                "Type",
                legend=None,
                scale=alt.Scale(domain=["Actual", "Expected"], range=[ACTUAL_COLOR, EXPECTED_COLOR]),
            ),
            tooltip=[
                alt.Tooltip("Type:N"),
                alt.Tooltip("Count:Q", title="Wins", format=".1~f"),
            ],
        )
        .properties(width=SMALL_CHART_W, height=SMALL_CHART_H, title=title)
    )


# ──────────────────────────────────────────────
# Win / loss chart
# ──────────────────────────────────────────────

def make_win_loss_chart(input_df: pd.DataFrame, input_player: str) -> alt.Chart:
    """Bar chart: wins and losses for a player (or all players)."""
    if input_player == "ALL PLAYERS":
        wins = input_df[input_df["Voittaja"] == input_df["Pelaaja vahvempi"]].shape[0]
    else:
        wins = input_df[input_df["Voittaja"] == input_player].shape[0]

    total = input_df[(input_df["Pelaaja vahvempi"] == input_player) |
                     (input_df["Pelaaja heikompi"] == input_player)].shape[0] if input_player != "ALL PLAYERS" \
        else input_df.shape[0]
    losses = total - wins

    data = pd.DataFrame({"Result": ["Wins", "Losses"], "Count": [wins, losses]})
    return _win_loss_bar(data)


# ──────────────────────────────────────────────
# Expected vs actual wins chart
# ──────────────────────────────────────────────

def make_expected_vs_actual_chart(filtered_df: pd.DataFrame, selected_player: str) -> alt.Chart:
    """Bar chart comparing expected win probability vs actual wins."""
    if selected_player == "ALL PLAYERS":
        expected_wins = filtered_df["Vahvemman voiton todennäköisyys"].sum()
        actual_wins = filtered_df[filtered_df["Voittaja"] == filtered_df["Pelaaja vahvempi"]].shape[0]
    else:
        expected_wins = filtered_df["Selected Player Win Probability"].sum()
        actual_wins = filtered_df[filtered_df["Voittaja"] == selected_player].shape[0]

    data = pd.DataFrame({"Type": ["Expected", "Actual"], "Count": [expected_wins, int(actual_wins)]})
    return _expected_vs_actual_bar(data)


# ──────────────────────────────────────────────
# Activity / performance timeline
# ──────────────────────────────────────────────

WEEKDAY_COLORS = {
    0: "#FFFFFF",  # Monday    – White
    1: "#FF4444",  # Tuesday   – Red
    2: "#44FF44",  # Wednesday – Green
    3: "#FFFF44",  # Thursday  – Yellow
    4: "#FF88FF",  # Friday    – Pink
    5: "#AA44FF",  # Saturday  – Purple
    6: "#FFAA44",  # Sunday    – Orange
}


def _adjust_color_intensity(base_color: str, game_count: int) -> str:
    """Scale a hex colour slightly brighter with more games (max ±20%)."""
    r, g, b = int(base_color[1:3], 16), int(base_color[3:5], 16), int(base_color[5:7], 16)
    intensity = 0.8 + (min(game_count / 5, 1) * 0.2)
    if base_color == "#FFFFFF":
        v = int(255 * intensity)
        return f"rgb({v},{v},{v})"
    return f"rgb({int(r*intensity)},{int(g*intensity)},{int(b*intensity)})"


def make_performance_chart(input_df: pd.DataFrame, input_player: str) -> alt.Chart:
    """Bar chart showing game activity over time, coloured by weekday."""
    if input_player == "ALL PLAYERS":
        melted = input_df.melt(
            id_vars=["Päivämäärä"],
            value_vars=["Pelaaja vahvempi", "Pelaaja heikompi"],
            var_name="Role",
            value_name="Player",
        )
        grouped = melted.groupby(["Päivämäärä", "Player"]).size().reset_index(name="Game Count")
        hover_field = "Player:N"
    else:
        mask = (input_df["Pelaaja vahvempi"] == input_player) | (input_df["Pelaaja heikompi"] == input_player)
        filtered = input_df[mask].copy()
        filtered["Opponent"] = filtered.apply(
            lambda r: r["Pelaaja heikompi"] if r["Pelaaja vahvempi"] == input_player else r["Pelaaja vahvempi"],
            axis=1,
        )
        grouped = filtered.groupby(["Päivämäärä", "Opponent"]).size().reset_index(name="Game Count")
        hover_field = "Opponent:N"

    grouped["Weekday"] = pd.to_datetime(grouped["Päivämäärä"]).dt.dayofweek
    grouped["Weekday Name"] = pd.to_datetime(grouped["Päivämäärä"]).dt.day_name()
    grouped["Color"] = grouped.apply(
        lambda r: _adjust_color_intensity(WEEKDAY_COLORS[r["Weekday"]], r["Game Count"]), axis=1
    )

    return (
        alt.Chart(grouped)
        .mark_bar()
        .encode(
            x=alt.X("Päivämäärä:T", title="Date"),
            y=alt.Y("Game Count:Q", title="Number of players"),
            color=alt.Color("Color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Päivämäärä:T", title="Date"),
                alt.Tooltip(hover_field, title="Player" if input_player == "ALL PLAYERS" else "Opponent"),
                alt.Tooltip("Game Count:Q", title="Games"),
                alt.Tooltip("Weekday Name:N", title="Weekday"),
            ],
        )
        .properties(width=800, height=LARGE_CHART_H, title=f"{input_player}'s club activity timeline")
        .interactive()
    )


# ──────────────────────────────────────────────
# Rating / rank timeline
# ──────────────────────────────────────────────

def _rating_to_rank_label(rating: float) -> str:
    """Convert a numeric ELO-style rating to a Go rank label (e.g. '1d', '5k')."""
    if rating >= 2100:
        return f"{int((rating - 2000) // 100)}d"
    return f"{int((2100 - rating) // 100)}k"


def make_rating_timeline_chart(
    input_df: pd.DataFrame, input_player: str, selected_opponent: str = "NONE"
) -> alt.Chart | None:
    """Line + point chart of rating over time, with rank labels on the y-axis.

    Draws an extra highlighted point at the most recent game to make the
    current rating immediately visible.
    """
    if input_player == "ALL PLAYERS":
        return None

    def _collect_ratings(df: pd.DataFrame, player: str) -> list[dict]:
        rows = []
        for col_role, col_rating in [("Pelaaja vahvempi", "Rating vahv"), ("Pelaaja heikompi", "Rating heik")]:
            for _, row in df[df[col_role] == player].iterrows():
                rows.append({"Date": row["Päivämäärä"], "Rating": row[col_rating], "Player": player})
        return rows

    ratings_p1 = _collect_ratings(input_df, input_player)
    ratings_p2 = _collect_ratings(input_df, selected_opponent) if selected_opponent != "NONE" else []

    if not ratings_p1:
        return None

    df_p1 = pd.DataFrame(ratings_p1).sort_values("Date")
    all_frames = [df_p1]
    if ratings_p2:
        df_p2 = pd.DataFrame(ratings_p2).sort_values("Date")
        all_frames.append(df_p2)

    combined = pd.concat(all_frames, ignore_index=True)

    # Y-axis range with padding
    r_min, r_max = combined["Rating"].min(), combined["Rating"].max()
    pad = (r_max - r_min) * 0.1 if r_max > r_min else 10
    y_min, y_max = r_min - pad, r_max + pad

    # Rating tick marks (every 100 points)
    ticks = list(range(int(y_min // 100) * 100, int(y_max // 100 + 1) * 100, 100))
    if not ticks:
        ticks = [2000, 2100] if combined["Rating"].nunique() > 1 else [int(combined["Rating"].iloc[0])]

    rank_df = pd.DataFrame({"value": ticks})
    rank_df["rank"] = rank_df["value"].apply(_rating_to_rank_label)

    y_scale = alt.Scale(domain=[y_min, y_max], nice=False)

    base = alt.Chart(combined).encode(x=alt.X("Date:T", title="Date"))

    line = base.mark_line().encode(
        y=alt.Y("Rating:Q", scale=y_scale, axis=alt.Axis(title="Rating", values=ticks, grid=True)),
        color=alt.Color("Player:N"),
        tooltip=[
            alt.Tooltip("Date:T", title="Date"),
            alt.Tooltip("Rating:Q", format=".0f", title="Rating"),
            alt.Tooltip("Player:N"),
        ],
    )

    points = base.mark_point(size=50).encode(
        y=alt.Y("Rating:Q", scale=y_scale),
        color=alt.Color("Player:N"),
        tooltip=[
            alt.Tooltip("Date:T", title="Date"),
            alt.Tooltip("Rating:Q", format=".0f", title="Rating"),
            alt.Tooltip("Player:N"),
        ],
    )

    # ── Latest-point highlight ──────────────────────────────────────
    latest_rows = []
    for player_name in combined["Player"].unique():
        player_df = combined[combined["Player"] == player_name]
        latest_rows.append(player_df.loc[player_df["Date"].idxmax()])
    latest_df = pd.DataFrame(latest_rows)

    latest_point = (
        alt.Chart(latest_df)
        .mark_point(size=180, shape="diamond", filled=True, stroke="white", strokeWidth=1.5)
        .encode(
            x=alt.X("Date:T"),
            y=alt.Y("Rating:Q", scale=y_scale),
            color=alt.Color("Player:N"),
            tooltip=[
                alt.Tooltip("Date:T", title="Latest game"),
                alt.Tooltip("Rating:Q", format=".0f", title="Current rating"),
                alt.Tooltip("Player:N"),
            ],
        )
    )

    latest_label = (
        alt.Chart(latest_df)
        .mark_text(align="right", dx=-10, dy=-8, fontWeight="bold")
        .encode(
            x=alt.X("Date:T"),
            y=alt.Y("Rating:Q", scale=y_scale),
            text=alt.Text("Rating:Q", format=".0f"),
            color=alt.Color("Player:N"),
        )
    )
    # ───────────────────────────────────────────────────────────────

    rating_labels = (
        alt.Chart(rank_df)
        .mark_text(align="right", baseline="middle", dx=-60)
        .encode(
            y=alt.Y("value:Q", scale=y_scale),
            text=alt.Text("value:Q", format=".0f"),
            color=alt.Color("rank:N", scale=alt.Scale(scheme="category10"), legend=None),
        )
    )

    rank_axis = (
        alt.Chart(rank_df)
        .mark_text(align="right", baseline="middle", dx=60, fontWeight="bold")
        .encode(
            y=alt.Y(
                "value:Q",
                scale=y_scale,
                axis=alt.Axis(orient="right", title="", values=ticks, grid=False),
            ),
            text="rank:N",
            color=alt.Color("rank:N", scale=alt.Scale(scheme="category10"), legend=None),
        )
    )

    title = f"{input_player}'s Rating Timeline"
    if selected_opponent != "NONE":
        title += f" vs {selected_opponent}"

    return (
        (line + points + latest_point + latest_label + rating_labels + rank_axis)
        .properties(width=800, height=LARGE_CHART_H, title=title)
        .interactive()
    )


# ──────────────────────────────────────────────
# Head-to-head charts
# ──────────────────────────────────────────────

def _filter_h2h(input_df: pd.DataFrame, player1: str, player2: str) -> pd.DataFrame:
    """Return only rows where player1 and player2 faced each other."""
    return input_df[
        ((input_df["Pelaaja vahvempi"] == player1) & (input_df["Pelaaja heikompi"] == player2))
        | ((input_df["Pelaaja vahvempi"] == player2) & (input_df["Pelaaja heikompi"] == player1))
    ]


def make_head_to_head_win_loss_chart(input_df: pd.DataFrame, player1: str, player2: str) -> alt.Chart:
    """Bar chart: wins and losses for player1 against player2."""
    h2h = _filter_h2h(input_df, player1, player2)
    wins = h2h[h2h["Voittaja"] == player1].shape[0]
    losses = h2h.shape[0] - wins
    data = pd.DataFrame({"Result": ["Wins", "Losses"], "Count": [wins, losses]})
    return _win_loss_bar(data, title=f"Games vs {player2}")


def make_head_to_head_expected_vs_actual_chart(input_df: pd.DataFrame, player1: str, player2: str) -> alt.Chart:
    """Bar chart comparing expected vs actual wins for player1 vs player2."""
    h2h = _filter_h2h(input_df, player1, player2)
    expected_wins = h2h.apply(
        lambda r: r["Vahvemman voiton todennäköisyys"]
        if r["Pelaaja vahvempi"] == player1
        else 1 - r["Vahvemman voiton todennäköisyys"],
        axis=1,
    ).sum()
    actual_wins = h2h[h2h["Voittaja"] == player1].shape[0]
    data = pd.DataFrame({"Type": ["Expected", "Actual"], "Count": [expected_wins, int(actual_wins)]})
    return _expected_vs_actual_bar(data, title=f"Wins vs {player2}")
