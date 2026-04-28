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


# ══════════════════════════════════════════════════════════════════
# CONFIGURATION  ← edit this block to add/remove seasons
# ══════════════════════════════════════════════════════════════════

SEASONS: list[dict] = [
    {
        "number": 1,
        "local_path": "data/goseason1.xlsx",
        "online_url": (
            "https://docs.google.com/spreadsheets/d/"
            "1NImVfaJ3z7K_hrlvFsC-HmmtXIj7x2ZVcIZthPf2IYs/export?format=xlsx"
        ),
        "usecols": "B:E,G,P,Q,W",
    },
    {
        "number": 2,
        "local_path": "data/goseason2.xlsx",
        "online_url": (
            "https://docs.google.com/spreadsheets/d/"
            "18qYVirc-_ni1I6myhICqA3hMgxP9pfvuBNJ0oLbljIE/export?format=xlsx"
        ),
        "usecols": "B:F,O,P,V",
    },
    {
        "number": 3,
        "local_path": "data/goseason3.xlsx",
        "online_url": (
            "https://docs.google.com/spreadsheets/d/"
            "1ktWkll-JubPHH3CSbcqO22WaJci2yAaQSK7-UKG8do8/export?format=xlsx"
        ),
        "usecols": "B:F,O,P,V",
    },
    {
        "number": 4,
        "local_path": "data/goseason4.xlsx",
        "online_url": (
            "https://docs.google.com/spreadsheets/d/"
            "1Mqn6RuWFyJTktBnJCfhOPngO8BUvB_K8rS_M8NdTSSE/export?format=xlsx"
        ),
        "usecols": "B:F,O,P,V",
    },
    {
        "number": 5,
        # Season 5 uses a special refresh-checked local file
        "local_path": "data/latest-season.xlsx",
        "online_url": (
            "https://docs.google.com/spreadsheets/d/"
            "1IjXQJSGUma8Deer0gT5uDDi0-VaDP3YFemMzHV1HLmA/export?format=xlsx"
        ),
        "usecols": "B:F,O,P,V",
        "is_latest": True,  # Triggers the freshness-check / auto-download logic
    },
]

# How long a cached local file is considered fresh before re-downloading
LATEST_SEASON_STALE_AFTER = timedelta(days=3)

# Canonical column names used throughout the app
COL_STRONGER = "Pelaaja vahvempi"
COL_WEAKER   = "Pelaaja heikompi"
COL_HANDICAP = "Tasoituskivet"
COL_WINNER   = "Voittaja"
COL_DATE     = "Päivämäärä"
COL_RATING_S = "Rating vahv"
COL_RATING_W = "Rating heik"
COL_WIN_PROB = "Vahvemman voiton todennäköisyys"

CANONICAL_COLS = [
    COL_STRONGER, COL_WEAKER, COL_HANDICAP, COL_WINNER,
    COL_DATE, COL_RATING_S, COL_RATING_W, COL_WIN_PROB,
]

# ══════════════════════════════════════════════════════════════════
# Page configuration
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Go club games dashboard",
    page_icon="🌑",
    layout="wide",
    initial_sidebar_state="expanded",
)
alt.theme.enable("dark")


# ══════════════════════════════════════════════════════════════════
# Data loading
# ══════════════════════════════════════════════════════════════════

def _download_file(url: str, local_path: str) -> None:
    """Download a file from *url* and save it to *local_path*."""
    response = requests.get(url)
    response.raise_for_status()
    with open(local_path, "wb") as fh:
        fh.write(response.content)


def _ensure_latest_season_fresh(cfg: dict) -> None:
    """Re-download the latest-season file if it is stale or missing."""
    local = cfg["local_path"]
    url   = cfg["online_url"]
    if os.path.exists(local):
        try:
            age = datetime.now() - datetime.fromtimestamp(os.path.getmtime(local))
            if age > LATEST_SEASON_STALE_AFTER:
                print(f"Local file is older than {LATEST_SEASON_STALE_AFTER}. Downloading…")
                _download_file(url, local)
        except Exception as exc:
            print(f"Error checking online file: {exc}. Using the local file.")
    else:
        print("Local file not found. Downloading…")
        _download_file(url, local)


@st.cache_data(ttl=timedelta(hours=6), max_entries=1)
def _load_season(cfg: dict) -> pd.DataFrame:
    """Load, clean and return a single season's DataFrame.

    Tries the local file first; falls back to the online URL.
    Returns an empty DataFrame if both fail.
    """
    read_kwargs = dict(
        engine="openpyxl",
        sheet_name="Pelitulokset",
        usecols=cfg["usecols"],
        skiprows=3,
    )
    for source_label, io in [("local", cfg["local_path"]), ("remote", cfg["online_url"])]:
        try:
            df = pd.read_excel(io=io, **read_kwargs)
            break
        except Exception as exc:
            msg = f"Season {cfg['number']}: failed to load from {source_label} ({exc})."
            if source_label == "local":
                st.warning(msg + " Trying remote…")
            else:
                st.error(msg)
                return pd.DataFrame()

    df = df.set_axis(CANONICAL_COLS, axis=1)
    df.dropna(axis=0, inplace=True)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], format="mixed", dayfirst=True)
    return df


def load_all_seasons() -> pd.DataFrame:
    """Load every season from SEASONS config, concat, and sort."""
    frames = []
    for cfg in SEASONS:
        if cfg.get("is_latest"):
            _ensure_latest_season_fresh(cfg)
        frames.append(_load_season(cfg))

    df = pd.concat(frames, ignore_index=True)
    df.sort_values(by=[COL_DATE, COL_STRONGER, COL_WEAKER], ascending=True, inplace=True)
    return df


df = load_all_seasons()


# ══════════════════════════════════════════════════════════════════
# Sidebar – filters
# ══════════════════════════════════════════════════════════════════

with st.sidebar:
    st.title("Go club games dashboard")

    # ── Player selector ──────────────────────────────────────────
    all_players = sorted(set(df[COL_STRONGER].unique()) | set(df[COL_WEAKER].unique()))
    all_players.insert(0, "ALL PLAYERS")
    selected_player = st.selectbox("Select player", all_players, index=0)

    # ── Opponent selector (only players who faced selected_player) ─
    if selected_player != "ALL PLAYERS":
        player_rows = df[(df[COL_STRONGER] == selected_player) | (df[COL_WEAKER] == selected_player)]
        opponents = set()
        for _, row in player_rows.iterrows():
            opponents.add(row[COL_WEAKER] if row[COL_STRONGER] == selected_player else row[COL_STRONGER])
        available_opponents = sorted(opponents)
    else:
        available_opponents = []

    available_opponents.insert(0, "NONE")
    selected_opponent = st.selectbox("Select opponent", available_opponents, index=0)

    # ── Date range ───────────────────────────────────────────────
    min_date, max_date = df[COL_DATE].min(), df[COL_DATE].max()
    selected_date_range = st.date_input("Select date range", [min_date, max_date])

    # ── Filter data ──────────────────────────────────────────────
    filtered_df = df[
        (df[COL_DATE] >= pd.to_datetime(selected_date_range[0]))
        & (df[COL_DATE] <= pd.to_datetime(selected_date_range[1]))
    ]
    if selected_player != "ALL PLAYERS":
        filtered_df = filtered_df[
            (filtered_df[COL_STRONGER] == selected_player)
            | (filtered_df[COL_WEAKER] == selected_player)
        ]

    # Weekday column
    filtered_df = filtered_df.copy()
    filtered_df["Weekday"] = filtered_df[COL_DATE].dt.day_name()

    # Per-player win probability
    if selected_player and selected_player != "ALL PLAYERS":
        filtered_df["Selected Player Win Probability"] = filtered_df.apply(
            lambda r: r[COL_WIN_PROB]
            if r[COL_STRONGER] == selected_player
            else 1 - r[COL_WIN_PROB],
            axis=1,
        )
    else:
        filtered_df["Selected Player Win Probability"] = None

    # ── Game details: optionally narrow to head-to-head ─────────
    game_details_df = filtered_df.copy()
    if selected_opponent != "NONE":
        game_details_df = game_details_df[
            ((game_details_df[COL_STRONGER] == selected_player) & (game_details_df[COL_WEAKER] == selected_opponent))
            | ((game_details_df[COL_STRONGER] == selected_opponent) & (game_details_df[COL_WEAKER] == selected_player))
        ]


# ══════════════════════════════════════════════════════════════════
# Dashboard top  (timeline + rating chart | win/loss | expected)
# ══════════════════════════════════════════════════════════════════

col = st.columns((8, 1.5, 1.5), gap="medium")

with col[0]:
    st.altair_chart(make_performance_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != "ALL PLAYERS":
        rating_chart = make_rating_timeline_chart(filtered_df, selected_player, selected_opponent)
        if rating_chart:
            st.altair_chart(rating_chart, use_container_width=True)

with col[1]:
    st.markdown("#### All games")
    st.altair_chart(make_win_loss_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != "ALL PLAYERS" and selected_opponent != "NONE":
        st.altair_chart(
            make_head_to_head_win_loss_chart(filtered_df, selected_player, selected_opponent),
            use_container_width=True,
        )

with col[2]:
    st.markdown("#### All wins")
    st.altair_chart(make_expected_vs_actual_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != "ALL PLAYERS" and selected_opponent != "NONE":
        st.altair_chart(
            make_head_to_head_expected_vs_actual_chart(filtered_df, selected_player, selected_opponent),
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════
# Dashboard main – Game details table
# ══════════════════════════════════════════════════════════════════

with st.container():
    st.markdown("#### Game details")

    sorted_df = game_details_df.sort_values(by=COL_DATE, ascending=False).copy()

    # ── Gor change column ───────────────────────────────────────
    # For each row, show how the selected player's rating changed vs the previous game.
    # For ALL PLAYERS view: show the stronger player's rating change.
    if selected_player != "ALL PLAYERS":
        player_games = sorted_df[
            (sorted_df[COL_STRONGER] == selected_player)
            | (sorted_df[COL_WEAKER] == selected_player)
        ].copy()
        player_games["_player_rating"] = player_games.apply(
            lambda r: r[COL_RATING_S] if r[COL_STRONGER] == selected_player else r[COL_RATING_W],
            axis=1,
        )
        player_games["Gor Change"] = player_games["_player_rating"].diff(-1)
        sorted_df = sorted_df.join(player_games[["Gor Change"]], how="left")
    else:
        sorted_df["Gor Change"] = sorted_df[COL_RATING_S].diff(-1)

    # Column order for display
    column_order = (
        COL_DATE, "Weekday", COL_RATING_S, COL_STRONGER, COL_WIN_PROB,
        COL_RATING_W, COL_WEAKER, COL_HANDICAP, COL_WINNER, "Gor Change",
    )

    st.dataframe(
        sorted_df,
        hide_index=False,
        column_order=column_order,
        column_config={
            COL_STRONGER: "Player (stronger)",
            COL_WIN_PROB: st.column_config.ProgressColumn(
                "Stronger player's win %",
                format="%.2f",
                help="Win probability for the stronger-rated player, based on rating difference and handicap.",
            ),
            COL_WEAKER: "Player (weaker)",
            COL_HANDICAP: st.column_config.NumberColumn(
                "Handicap stones",
                format="%d",
                help="Number of handicap stones given to the weaker player.",
            ),
            COL_WINNER: "Winner",
            COL_DATE: st.column_config.DateColumn(
                "Date",
                format="YYYY-MM-DD",
                help="Date the club game was played.",
            ),
            COL_RATING_S: st.column_config.NumberColumn(
                "Rating (stronger)",
                format="%.0f",
                help="Club ELO-style rating of the stronger player before this game.",
            ),
            COL_RATING_W: st.column_config.NumberColumn(
                "Rating (weaker)",
                format="%.0f",
                help="Club ELO-style rating of the weaker player before this game.",
            ),
            "Gor Change": st.column_config.NumberColumn(
                "Gor Δ",
                format="%+.0f",
                help=(
                    "Rating change for the selected player (or stronger player in ALL PLAYERS view) "
                    "from the previous game to this one. Positive = rating increased."
                ),
            ),
        },
    )


# ══════════════════════════════════════════════════════════════════
# Bottom info
# ══════════════════════════════════════════════════════════════════

with st.expander("About the stats for the selected player", expanded=False):
    st.write("""
- **Player's club games timeline**: Activity in club games over time. Bars are coloured by weekday.
- **Club games timeline (ALL PLAYERS)**: Counts both players per game (total = games × 2).
- **Games**: Wins and losses. ALL PLAYERS view uses the higher-rated player of each game.
- **Wins**: Expected wins (from ratings + handicap) vs actual wins.
- **Wins (ALL PLAYERS)**: Based on the stronger-rated player of each game.
- **Player's rating timeline**: Club Gor rating over time. The diamond ◆ marks the latest game.
- **Game details**: All recorded club games with statistics, including rating change per game (Gor Δ).
""")

with st.expander("Update history", expanded=False):
    st.write("""
#### Updates 2025 (current):
1. **Gor Δ column**: Rating change per game now shown in the game details table.
2. **Latest-point marker**: The rating graph now draws a highlighted diamond at the most recent game.
3. **Hover tooltips**: All charts now show richer hover information (weekday, opponent, formatted values).
4. **Season config block**: Adding a new season now requires editing only the SEASONS list at the top of the file.
5. **Code clarity**: Shared chart helpers, named constants for columns, cleaner separation of concerns.

#### Updates 21.9.2025:
1. **4th season added**: New club game season began on September 8th.
2. **Updated about section**: To include all parts of the dashboard.
3. **New colours based on weekdays**: Shows each weekday as a specific colour.

#### Updates 26.5.2025:
1. **Opponent selection**: Shows only the games between players.
2. **Performance optimization**: Caches the data.
3. **Face-to-face comparison**: Compare the rating timelines of two players.
4. **Face-to-face wins**: Bar charts for head-to-head comparison.

#### Updates 23.5.2025:
1. **Rating graph**: shows player's rating over time.
2. **Download optimization**: checks if the local file is older than 3 days before downloading.
3. **Included all players to timeline**: not just half of them.
4. **Ordered game details**: newest first.
5. **Decimal formatting**: for expected wins and rating.

#### Updates 15.5.2025:
1. **Timeline update**: shows opponents.
2. **Added expected wins**: and comparison to actual wins.
3. **Rearranged game details**: win probability visualized and column order improved.
4. **Added statistics for ALL PLAYERS**.

#### Updates 14.5.2025:
1. **Added third season games**.
2. **Included more data**: player's ratings, expected win %.
3. **Updated timeline**: Game colour by winner, barchart based on played game dates.

#### Prototype 17.2.2025:
1. **Data Loading from online spreadsheets**.
2. **Sidebar**: player and date range filter.
3. **Timeline**, **Win/Loss Ratio**, **Recent games**.

#### Next steps:
- **Handicap analysis**: How players perform with handicap stones.
- **Confidence intervals**: Error margins based on games played so far.
- **Translation**: Finnish / English toggle.
- **Player selection from game details table**.
- **All players' timeline**: Player colour based on amount of games.
- **Accessibility testing**: Colour-blindness, contrast, alt-texts, keyboard navigation.
- **Visual look**: Go art, board, stones, cups.
- **Opponent rating graph**: Show both players' full history when opponent is selected.
- **Head-to-head summary stats**: Win %, average Gor change, handicap breakdown.
""")
