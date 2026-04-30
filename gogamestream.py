import glob
import os
import warnings

import altair as alt
import pandas as pd
import streamlit as st

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
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════

DATA_DIR    = "data"
LATEST_FILE = os.path.join(DATA_DIR, "latest-season.xlsx")

# Columns to read from the Excel sheets.
# Season 1 has a different layout from all later seasons.
# Base cols: players, handicap, winner, date, ratings, win probability.
# Z,AE (S1) / Y,AD (S2+): Pelaaja 1 & 2 rating muutos (Gor change).
USECOLS_S1 = "B:E,G,P,Q,W,Z,AE"
USECOLS    = "B:F,O,P,V,Y,AD"

# Canonical column names used throughout the app
COL_STRONGER = "Pelaaja vahvempi"
COL_WEAKER   = "Pelaaja heikompi"
COL_HANDICAP = "Tasoituskivet"
COL_WINNER   = "Voittaja"
COL_DATE     = "Päivämäärä"
COL_RATING_S = "Rating vahv"
COL_RATING_W = "Rating heik"
COL_WIN_PROB = "Vahvemman voiton todennäköisyys"

COL_GOR_S    = "Gor Δ (stronger)"
COL_GOR_W    = "Gor Δ (weaker)"

CANONICAL_COLS = [
    COL_STRONGER, COL_WEAKER, COL_HANDICAP, COL_WINNER,
    COL_DATE, COL_RATING_S, COL_RATING_W, COL_WIN_PROB,
    COL_GOR_S, COL_GOR_W,
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
# Data loading  (downloading is handled by update-data.yml)
# ══════════════════════════════════════════════════════════════════

def _discover_season_files() -> list[tuple[int, str]]:
    """Return (season_number, path) pairs for every season file, sorted by number.

    Scans DATA_DIR for goseason{N}.xlsx files plus latest-season.xlsx.
    Season number for latest-season is one above the highest archived season.
    """
    archived = sorted(
        glob.glob(os.path.join(DATA_DIR, "goseason[0-9]*.xlsx"))
    )
    season_nums = []
    for path in archived:
        basename = os.path.basename(path)
        try:
            num = int(basename.replace("goseason", "").replace(".xlsx", ""))
            season_nums.append((num, path))
        except ValueError:
            pass

    if os.path.exists(LATEST_FILE):
        latest_num = (max(n for n, _ in season_nums) + 1) if season_nums else 1
        season_nums.append((latest_num, LATEST_FILE))

    return season_nums


def _usecols_for(season_number: int) -> str:
    """Return the correct usecols string for this season number."""
    return USECOLS_S1 if season_number == 1 else USECOLS


@st.cache_data(show_spinner="Loading season data…")
def load_all_seasons(file_mtimes: dict[str, float]) -> pd.DataFrame:
    """Load, clean, concat and sort all season files into one DataFrame.

    *file_mtimes* is passed purely so Streamlit re-runs this function whenever
    any data file changes on disk (mtime changes → cache miss).
    """
    read_kwargs = dict(engine="openpyxl", sheet_name="Pelitulokset", skiprows=3)
    frames = []

    for season_num, path in _discover_season_files():
        try:
            raw = pd.read_excel(io=path, usecols=_usecols_for(season_num), **read_kwargs)
        except Exception as exc:
            st.error(f"Season {season_num}: could not load '{path}': {exc}")
            continue

        raw = raw.set_axis(CANONICAL_COLS, axis=1)
        raw.dropna(axis=0, inplace=True)
        raw[COL_DATE] = pd.to_datetime(raw[COL_DATE], format="mixed", dayfirst=True)
        frames.append(raw)

    if not frames:
        st.error("No season data found in the data/ directory.")
        return pd.DataFrame(columns=CANONICAL_COLS)

    combined = pd.concat(frames, ignore_index=True)
    combined.sort_values(by=[COL_DATE, COL_STRONGER, COL_WEAKER], ascending=True, inplace=True)
    return combined


def _current_mtimes() -> dict[str, float]:
    """Snapshot the mtime of every season file so cache can detect changes."""
    return {
        path: os.path.getmtime(path)
        for _, path in _discover_season_files()
        if os.path.exists(path)
    }


df = load_all_seasons(_current_mtimes())


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
    all_dates = sorted(df[COL_DATE].dt.date.unique())
    min_date, max_date = all_dates[0], all_dates[-1]

    from_date = st.selectbox(
        "From date",
        options=all_dates,
        index=0,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
    # Default "To date" to the end of the data, but allow independent selection.
    to_date = st.selectbox(
        "To date",
        options=all_dates,
        index=len(all_dates) - 1,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
    if to_date < from_date:
        st.warning("'To date' is before 'From date' — no data will be shown.")

    # ── Filter data ──────────────────────────────────────────────
    filtered_df = df[
        (df[COL_DATE] >= pd.to_datetime(from_date))
        & (df[COL_DATE] <= pd.to_datetime(to_date))
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
        # Pass the full df (date-filtered only) so both players' complete rating
        # histories are shown, not just games they played against each other.
        date_filtered_df = df[
            (df[COL_DATE] >= pd.to_datetime(from_date))
            & (df[COL_DATE] <= pd.to_datetime(to_date))
        ]
        rating_chart = make_rating_timeline_chart(date_filtered_df, selected_player, selected_opponent)
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

    # Column order: Date | Weekday | Rating(s) | GorΔ(s) | Player(s) | Win% | Rating(w) | GorΔ(w) | Player(w) | Handicap | Winner
    column_order = (
        COL_DATE, "Weekday",
        COL_RATING_S, COL_GOR_S, COL_STRONGER, COL_WIN_PROB,
        COL_RATING_W, COL_GOR_W, COL_WEAKER,
        COL_HANDICAP, COL_WINNER,
    )

    st.dataframe(
        sorted_df,
        hide_index=False,
        column_order=column_order,
        column_config={
            COL_DATE: st.column_config.DateColumn(
                "Date",
                format="YYYY-MM-DD",
                help="Date the club game was played.",
            ),
            COL_RATING_S: st.column_config.NumberColumn(
                "Rating (stronger)",
                format="%.0f",
                help="Club Gor rating of the stronger player before this game.",
            ),
            COL_GOR_S: st.column_config.NumberColumn(
                "Gor Δ (s.)",
                format="%+.1f",
                help="Rating change for the stronger player from this game.",
            ),
            COL_STRONGER: "Player (stronger)",
            COL_WIN_PROB: st.column_config.ProgressColumn(
                "Stronger player's win %",
                format="%.2f",
                help="Win probability for the stronger-rated player, based on rating difference and handicap.",
            ),
            COL_RATING_W: st.column_config.NumberColumn(
                "Rating (weaker)",
                format="%.0f",
                help="Club Gor rating of the weaker player before this game.",
            ),
            COL_GOR_W: st.column_config.NumberColumn(
                "Gor Δ (w.)",
                format="%+.1f",
                help="Rating change for the weaker player from this game.",
            ),
            COL_WEAKER: "Player (weaker)",
            COL_HANDICAP: st.column_config.NumberColumn(
                "Handicap stones",
                format="%d",
                help="Number of handicap stones given to the weaker player.",
            ),
            COL_WINNER: "Winner",
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
- **Player's rating timeline**: Club gor rating over time. Each point shows the rating **after** the game.
- **Game details**: All recorded club games with statistics, including rating change per game (Gor Δ).

**ATTENTION**: 

**New games after the season end date should be added to the new season's spreadsheet with the same URL as the old season (old seasons should be moved and archived.**
""")

with st.expander("Update history", expanded=False):
    st.write("""
#### Updates 28.4.2026:
1. **Gor change column**: Shows rating change per game in the game details table.
2. **Latest gor**: The rating graph now draws a highlighted diamond at the most recent game.
3. **Hover tooltips**: All charts show simple info.
4. **Seasons change automatically** when the season's spreadsheet date comes up.
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
