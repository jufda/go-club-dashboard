# 🌑 Go Club Dashboard

An interactive dashboard for analyzing Go club game statistics and player performance.

**[View the online dashboard →](https://go-games-dashboard.streamlit.app/)**

## Features

- 📊 **Player Statistics** – Win/loss records, expected vs. actual wins, and performance metrics
- 📈 **Rating Timeline** – Track player Gor ratings over time with visual progression
- 🎮 **Head-to-Head Analysis** – Compare matchup records and win probabilities between any two players
- 📅 **Activity Timeline** – View game frequency colored by weekday
- 🔍 **Flexible Filtering** – Analyze data by player, opponent, and date range
- 🧘 **Go Proverbs** – Philosophical Go wisdom displayed with each session
- ⚡ **Auto-updating** – Scheduled data fetches from Google Sheets (Mondays & Thursdays)

## Set up for your own go club

Want to track your club's games? Follow these steps:

### 1. Prepare your game data

Create a spreadsheet with the following columns:
- **Pelaaja vahvempi** (Stronger player's name)
- **Pelaaja heikompi** (Weaker player's name)  
- **Tasoituskivet** (Handicap stones)
- **Voittaja** (Winner)
- **Päivämäärä** (Date)
- **Rating vahv** (Stronger player's rating before game)
- **Rating heik** (Weaker player's rating before game)
- **Vahvemman voiton todennäköisyys** (Expected win probability for stronger player)
- **Gor Δ (stronger)** (Rating change for stronger player)
- **Gor Δ (weaker)** (Rating change for weaker player)

*Note: If your columns are in a different language, update the column name constants in `gogamestream.py`.*

### 2. Update the data URL

Edit [gogamestream.py](gogamestream.py) and replace the spreadsheet URL:

```python
URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=xlsx"
```

To get your sheet's ID:
- Open your [Proton Drive](https://proton.me/drive/docs) or Google Sheet
- Look at the URL:  `drive.proton.me/url/{ID}` or `docs.google.com/spreadsheets/d/{ID}/edit`
- Copy the ID.

### 3. Set up automatic updates

The repository includes a GitHub Actions workflow (`.github/workflows/update-data.yml`) that automatically downloads your latest game data. To enable it:

1. Push this repository to GitHub
2. Ensure Google Sheets export link is public (or adjust permissions as needed)
3. Adjust the workflow runs to local club schedule (the nights following the game days)
4. You can also trigger it manually from the Actions tab

### 4. Deploy to Streamlit cloud

1. Push your customized repository to GitHub
2. Sign up at [Streamlit cloud](https://share.streamlit.io/)
3. Create a new app, select your repository and `gogamestream.py`
4. Streamlit will automatically deploy and update your dashboard


## Feature ideas

[Share your ideas in the discussions!](../../discussions)


## Architecture

- **[gogamestream.py](gogamestream.py)** – Main Streamlit app and UI logic
- **[charts.py](charts.py)** – Altair chart generation and utilities
- **[go_proverbs.py](go_proverbs.py)** – Classical Go wisdom database
- **[data/](data/)** – Cached game data (auto-updated)

**Technologies:**
- [Streamlit](https://streamlit.io/) – Interactive web framework
- [Pandas](https://pandas.pydata.org/) – Data manipulation
- [Altair](https://altair-viz.github.io/) – Declarative visualization
- [openpyxl](https://openpyxl.readthedocs.io/) – Excel file reading
- [GitHub Actions](https://github.com/features/actions) – Automated data updates

## License

This project is licensed under the [GPLv3](LICENSE).

---

**Made with <3 for the Go community. Good luck and enjoy your games!**