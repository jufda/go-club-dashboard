import glob
import os
import warnings

import altair as alt
import pandas as pd
import streamlit as st

from charts import (
    make_expected_vs_actual_chart,
    make_head_to_head_expected_vs_actual_chart,
    make_head_to_head_win_loss_chart,
    make_performance_chart,
    make_rating_timeline_chart,
    make_win_loss_chart,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ══════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════

DATA_DIR    = "data"
LATEST_FILE = os.path.join(DATA_DIR, "latest-season.xlsx")

ALL_PLAYERS = "ALL PLAYERS"
NO_OPPONENT = "NONE"

# Columns to read from the Excel sheets.
# Season 1 has a different layout from all later seasons.
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
    """Return (season_number, path) pairs for every season file, sorted."""
    archived = sorted(glob.glob(os.path.join(DATA_DIR, "goseason[0-9]*.xlsx")))
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


def _current_mtimes() -> dict[str, float]:
    """Snapshot mtime of every season file so cache detects changes."""
    return {
        path: os.path.getmtime(path)
        for _, path in _discover_season_files()
        if os.path.exists(path)
    }


@st.cache_data(show_spinner="Loading season data…")
def load_all_seasons(file_mtimes: dict[str, float]) -> pd.DataFrame:
    """Load, clean, concat and sort all season files into one DataFrame."""
    read_kwargs = dict(engine="openpyxl", sheet_name="Pelitulokset", skiprows=3)
    frames = []

    for season_num, path in _discover_season_files():
        usecols = USECOLS_S1 if season_num == 1 else USECOLS
        try:
            raw = pd.read_excel(io=path, usecols=usecols, **read_kwargs)
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


df = load_all_seasons(_current_mtimes())
all_players = sorted(set(df[COL_STRONGER].unique()) | set(df[COL_WEAKER].unique()))


# ══════════════════════════════════════════════════════════════════
# Session state — player selection (supports click-from-table)
# ══════════════════════════════════════════════════════════════════

# A player name in ?player=Name (set by clicking a name cell in the table)
# takes priority over the existing session state value, then clears itself.
_qp_player = st.query_params.get("player", "")
if _qp_player in all_players:
    st.session_state["selected_player"] = _qp_player
    st.query_params.clear()

if "selected_player" not in st.session_state:
    st.session_state["selected_player"] = ALL_PLAYERS


# ══════════════════════════════════════════════════════════════════
# Sidebar – filters
# ══════════════════════════════════════════════════════════════════

with st.sidebar:
    st.title("Go club games dashboard")

    # ── Player selector ──────────────────────────────────────────
    player_options = [ALL_PLAYERS] + all_players
    player_index = (
        player_options.index(st.session_state["selected_player"])
        if st.session_state["selected_player"] in player_options
        else 0
    )
    selected_player = st.selectbox(
        "Select player", player_options, index=player_index, key="player_selectbox"
    )
    st.session_state["selected_player"] = selected_player

    # ── Opponent selector ────────────────────────────────────────
    if selected_player != ALL_PLAYERS:
        mask = (df[COL_STRONGER] == selected_player) | (df[COL_WEAKER] == selected_player)
        opponents = sorted(
            (df.loc[mask, COL_WEAKER].where(df[COL_STRONGER] == selected_player)
             .combine_first(df.loc[mask, COL_STRONGER])).dropna().unique()
        )
    else:
        opponents = []

    selected_opponent = st.selectbox("Select opponent", [NO_OPPONENT] + opponents, index=0)

    # ── Date range ───────────────────────────────────────────────
    all_dates = sorted(df[COL_DATE].dt.date.unique())
    from_date = st.selectbox(
        "From date", options=all_dates, index=0,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
    to_date = st.selectbox(
        "To date", options=all_dates, index=len(all_dates) - 1,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
    if to_date < from_date:
        st.warning("'To date' is before 'From date' — no data will be shown.")


# ══════════════════════════════════════════════════════════════════
# Filtered data
# ══════════════════════════════════════════════════════════════════

date_mask = (df[COL_DATE] >= pd.Timestamp(from_date)) & (df[COL_DATE] <= pd.Timestamp(to_date))
date_filtered_df = df[date_mask].copy()
date_filtered_df["Weekday"] = date_filtered_df[COL_DATE].dt.day_name()

if selected_player != ALL_PLAYERS:
    player_mask = (
        (date_filtered_df[COL_STRONGER] == selected_player)
        | (date_filtered_df[COL_WEAKER] == selected_player)
    )
    filtered_df = date_filtered_df[player_mask].copy()
    filtered_df["Selected Player Win Probability"] = filtered_df.apply(
        lambda r: r[COL_WIN_PROB] if r[COL_STRONGER] == selected_player else 1 - r[COL_WIN_PROB],
        axis=1,
    )
else:
    filtered_df = date_filtered_df.copy()
    filtered_df["Selected Player Win Probability"] = None

# Head-to-head narrowing for the game details table
game_details_df = filtered_df.copy()
if selected_opponent != NO_OPPONENT:
    h2h_mask = (
        ((game_details_df[COL_STRONGER] == selected_player) & (game_details_df[COL_WEAKER] == selected_opponent))
        | ((game_details_df[COL_STRONGER] == selected_opponent) & (game_details_df[COL_WEAKER] == selected_player))
    )
    game_details_df = game_details_df[h2h_mask]


# ══════════════════════════════════════════════════════════════════
# Dashboard top  (timeline | win/loss | expected)
# ══════════════════════════════════════════════════════════════════

col = st.columns((8, 1.5, 1.5), gap="medium")

with col[0]:
    st.altair_chart(make_performance_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != ALL_PLAYERS:
        rating_chart = make_rating_timeline_chart(date_filtered_df, selected_player, selected_opponent)
        if rating_chart:
            st.altair_chart(rating_chart, use_container_width=True)

with col[1]:
    st.markdown("#### All games")
    st.altair_chart(make_win_loss_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != ALL_PLAYERS and selected_opponent != NO_OPPONENT:
        st.altair_chart(
            make_head_to_head_win_loss_chart(filtered_df, selected_player, selected_opponent),
            use_container_width=True,
        )

with col[2]:
    st.markdown("#### All wins")
    st.altair_chart(make_expected_vs_actual_chart(filtered_df, selected_player), use_container_width=True)

    if selected_player != ALL_PLAYERS and selected_opponent != NO_OPPONENT:
        st.altair_chart(
            make_head_to_head_expected_vs_actual_chart(filtered_df, selected_player, selected_opponent),
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════
# Game details table  (click player name cell → selects that player)
# ══════════════════════════════════════════════════════════════════

with st.container():
    st.markdown("#### Game details")
    st.caption("💡 Click a player name to select them.")

    sorted_df = game_details_df.sort_values(by=COL_DATE, ascending=False).reset_index(drop=True)

    # Build ?player=Name links so clicking a name cell sets the player filter.
    # Using st.query_params: the app re-runs, reads the param, updates session
    # state, clears the param, and the sidebar selectbox updates automatically.
    def _player_url(name: str) -> str:
        return f"?player={name}"

    display_df = sorted_df.copy()
    display_df[COL_STRONGER] = display_df[COL_STRONGER].apply(_player_url)
    display_df[COL_WEAKER]   = display_df[COL_WEAKER].apply(_player_url)

    column_order = (
        COL_DATE, "Weekday",
        COL_RATING_S, COL_GOR_S, COL_STRONGER, COL_WIN_PROB,
        COL_RATING_W, COL_GOR_W, COL_WEAKER,
        COL_HANDICAP, COL_WINNER,
    )

    st.dataframe(
        display_df,
        hide_index=True,
        column_order=column_order,
        column_config={
            COL_DATE: st.column_config.DateColumn("Date", format="YYYY-MM-DD",
                help="Date the club game was played."),
            COL_RATING_S: st.column_config.NumberColumn("Rating (stronger)", format="%.0f",
                help="Club Gor rating of the stronger player before this game."),
            COL_GOR_S: st.column_config.NumberColumn("Gor Δ (s.)", format="%+.1f",
                help="Rating change for the stronger player from this game."),
            COL_STRONGER: st.column_config.LinkColumn(
                "Player (stronger)",
                display_text=r"([^?=]+)$",
                help="Click to select this player.",
            ),
            COL_WIN_PROB: st.column_config.ProgressColumn("Stronger win %", format="%.2f",
                help="Win probability for the stronger-rated player."),
            COL_RATING_W: st.column_config.NumberColumn("Rating (weaker)", format="%.0f",
                help="Club Gor rating of the weaker player before this game."),
            COL_GOR_W: st.column_config.NumberColumn("Gor Δ (w.)", format="%+.1f",
                help="Rating change for the weaker player from this game."),
            COL_WEAKER: st.column_config.LinkColumn(
                "Player (weaker)",
                display_text=r"([^?=]+)$",
                help="Click to select this player.",
            ),
            COL_HANDICAP: st.column_config.NumberColumn("Handicap stones", format="%d",
                help="Number of handicap stones given to the weaker player."),
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
- **Player's rating timeline**: Club Gor rating over time. The diamond ◆ marks the latest game.
- **Game details**: All recorded club games with statistics, including rating change per game (Gor Δ).
- **Click to select**: Click any row in the Game details table to reveal player buttons.
""")

with st.expander("Update history", expanded=False):
    st.write("""
#### Updates 2026:
1. **Click-to-select player**: Click any row in the game details table — player buttons appear instantly.
2. **UTC fix**: Workflow now uses timezone-aware `datetime.now(UTC)` instead of deprecated `utcnow()`.
3. **Code improvements**: Named sentinel constants (`ALL_PLAYERS`, `NO_OPPONENT`), cleaner filtering,
   session state for player selection, opponent list derived without iterrows.

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
2. **Included all players to timeline**: not just half of them.
3. **Ordered game details**: newest first.
4. **Decimal formatting**: for expected wins and rating.

#### Prototype 17.2.2025:
1. **Data Loading from online spreadsheets**.
2. **Sidebar**: player and date range filter.
3. **Timeline**, **Win/Loss Ratio**, **Recent games**.

#### Next steps:
- **Handicap analysis**: How players perform with handicap stones.
- **Confidence intervals**: Error margins based on games played so far.
- **Translation**: Finnish / English toggle.
- **All players' timeline**: Player colour based on amount of games.
- **Accessibility testing**: Colour-blindness, contrast, alt-texts, keyboard navigation.
- **Visual look**: Go art, board, stones, cups.
- **Head-to-head summary stats**: Win %, average Gor change, handicap breakdown.
""")
