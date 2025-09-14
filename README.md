# IMDb Movie Automation

Python automation to scrape movie data from IMDb and email results.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Create `.env` file:
```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your-email@gmail.com
SMTP_PASSWORD=your-app-password
EMAIL_TO=recipient@gmail.com
EMAIL_FROM=your-email@gmail.com
```

3. Create `movies.xlsx` with a "Movies" column containing movie titles.

## Usage

```bash
python main.py
```

## Output

- SQLite database: `output/movies.sqlite3`
- CSV export: `output/movies.csv`
- Email with attachments sent automatically

## Data Extracted

- IMDb rating
- Popularity score
- Metascore
- Genres
- Top 5 user reviews
- User review count