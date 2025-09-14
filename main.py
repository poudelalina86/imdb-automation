from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp
import sqlite3
import os

import re
from dotenv import load_dotenv

 
load_dotenv()

 
browser = Selenium()

 
def scrape_movies_and_email():
    """Read movies from Excel, scrape IMDb and Rotten Tomatoes, save to SQLite, and email results"""
    browser.open_available_browser("https://www.imdb.com", maximized=True)

    database_path = os.path.join("output", "movies.sqlite3")
    db = _init_db(database_path)

    movies = _read_movie_titles_from_excel("movies.xlsx")
    print(f"Movies to process: {movies}")

    for movie_title in movies:
        print(f"\n=== Processing: {movie_title} ===")
        try:
            imdb_match = _imdb_find_exact_movie(movie_title)
            if imdb_match is None:
                print(f"No exact match found for: {movie_title}")
                _insert_row(db, movie_title, None, None, None, None, None, [None]*5, status="No exact match found")
                continue

            imdb_url, year = imdb_match
            print(f"Found IMDb match: {imdb_url} (year: {year})")
            
            imdb_rating, popularity, metascore, genres, featured_reviews, user_reviews_count = _imdb_extract_details(imdb_url)
            print(f"IMDb data - Rating: {imdb_rating}, Popularity: {popularity}, Metascore: {metascore}, Genres: {genres}")
            
            # Convert featured_reviews string back to list for compatibility
            reviews_list = featured_reviews.split("\n---\n") if featured_reviews else [None]*5
            while len(reviews_list) < 5:
                reviews_list.append(None)

            _insert_row(
                db,
                movie_title,
                imdb_rating,
                popularity,
                metascore,
                user_reviews_count,
                genres,
                reviews_list[:5],
                status="success",
            )
            print(f"Successfully saved data for: {movie_title}")
        except Exception as e:
            print(f"Error processing {movie_title}: {e}")
            _insert_row(db, movie_title, None, None, None, None, None, [None]*5, status=f"error: {e}")

    db.commit()
    _email_results(database_path)
    browser.close_all_browsers()

 
def _init_db(db_path: str):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
     
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='movies'")
    if cur.fetchone():
        cur.execute("PRAGMA table_info(movies)")
        columns = [col[1] for col in cur.fetchall()]
        if 'movie_name' in columns:  # Old schema
            cur.execute("DROP TABLE movies")
    
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS movies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            rating TEXT,
            popularity TEXT,
            metascore TEXT,
            genre TEXT,
            featured_reviews TEXT,
            user_reviews TEXT,
            status TEXT
        )
        """
    )
    conn.commit()
    return conn

def _read_movie_titles_from_excel(path: str):
    excel = Files()
    excel.open_workbook(path)
    table = excel.read_worksheet_as_table(header=True)   
    excel.close_workbook()

    if not table:
        return []

 
    if isinstance(table[0], dict):
        col_name = None
        for h in table[0].keys():
            if str(h).strip().lower() == "movies":
                col_name = h
                break
        if not col_name:
            raise ValueError("Excel file must have a 'Movies' column")
        titles = [str(row[col_name]).strip() for row in table if row.get(col_name)]
    else:
        
        titles = []
        for row in table[1:]:  
            if row and row[0]:
                titles.append(str(row[0]).strip())

    return titles


def _imdb_find_exact_movie(query_title: str):
    """Find exact movie match using IMDb find page like Robot Framework"""
    query_encoded = query_title.replace(" ", "+")
    find_url = f"https://www.imdb.com/find/?q={query_encoded}&s=tt&ttype=ft"
    browser.go_to(find_url)
    
    browser.wait_until_element_is_visible("xpath://section[contains(@data-testid,'find-results-section-title')] | //table[contains(@class,'findList')]", timeout=10)
    
    exact_matches = []
    title_clean = query_title.strip().lower()
     
    try:
        results = browser.find_elements("xpath://section[contains(@data-testid,'find-results-section-title')]//li[contains(@class,'ipc-metadata-list-summary-item')]")
        for result in results:
            try:
                title_element = browser.find_element("xpath:.//a[contains(@class,'ipc-metadata-list-summary-item__t')]", parent=result)
                title = title_element.text.strip().lower()
                
                if title != title_clean:
                    continue
                
                # Check if it's a movie (not TV)
                try:
                    type_element = browser.find_element("xpath:.//span[contains(@class,'ipc-metadata-list-summary-item__tl')]", parent=result)
                    type_text = type_element.text.lower()
                    if "tv" in type_text:
                        continue
                except:
                    pass
                
                href = title_element.get_attribute("href")
                year_match = re.search(r'\d{4}', result.text)
                year = int(year_match.group(0)) if year_match else 0
                
                full_url = f"https://www.imdb.com{href}" if href.startswith("/") else href
                exact_matches.append((full_url, year))
            except Exception:
                continue
    except:
        pass
    
    # Try old UI if new UI failed
    if not exact_matches:
        try:
            results = browser.find_elements("xpath://table[contains(@class,'findList')]//tr")
            for result in results:
                try:
                    title_element = browser.find_element("xpath:.//td[@class='result_text']/a", parent=result)
                    title = title_element.text.strip().lower()
                    
                    if title != title_clean:
                        continue
                    
                    
                    meta_text = browser.get_text("xpath:.//td[@class='result_text']", parent=result).lower()
                    if "tv" in meta_text:
                        continue
                    
                    href = title_element.get_attribute("href")
                    year_match = re.search(r'\d{4}', meta_text)
                    year = int(year_match.group(0)) if year_match else 0
                    
                    full_url = f"https://www.imdb.com{href}" if href.startswith("/") else href
                    exact_matches.append((full_url, year))
                except Exception:
                    continue
        except:
            pass
    
    if not exact_matches:
        return None
    
    # Return most recent year
    exact_matches.sort(key=lambda x: x[1], reverse=True)
    return exact_matches[0]

def _imdb_extract_details(title_url: str):
    """Extract IMDb details matching Robot Framework approach"""
    browser.go_to(title_url)
    browser.wait_until_element_is_visible("xpath://span[@data-testid='hero__primary-text'] | //h1", timeout=10)
    
    # Extract IMDb rating
    imdb_rating = None
    try:
        imdb_rating = browser.get_text("xpath://div[@data-testid='hero-rating-bar__aggregate-rating__score']//span[1]")
    except:
        pass
    
    # Extract popularity
    popularity = None
    try:
        popularity = browser.get_text("xpath://div[@data-testid='hero-rating-bar__popularity__score']")
    except:
        pass
    
    # Extract metascore
    metascore = None
    try:
        metascore = browser.get_text("xpath://span[contains(@class,'metacritic-score-box')]")
    except:
        pass
    
    # Extract genres
    genres = []
    try:
        genre_elements = browser.find_elements("xpath://div[@data-testid='genres']//span[contains(@class,'ipc-chip__text')]")
        genres = [elem.text.strip() for elem in genre_elements]
    except:
        pass
    
    # Extract user reviews count
    user_reviews_count = None
    try:
        user_reviews_count = browser.get_text("xpath://a[contains(@href,'/reviews')][.//span[contains(.,'User reviews')]]//span[contains(@class,'score')]")
    except:
        pass
    
    # Get top 5 reviews from reviews page
    reviews_list = []
    title_match = re.search(r'/title/(tt\d+)', title_url)
    if title_match:
        ttid = title_match.group(1)
        reviews_url = f"https://www.imdb.com/title/{ttid}/reviews/?sort=helpfulnessScore,desc"
        try:
            browser.go_to(reviews_url)
            browser.wait_until_element_is_visible("xpath://*[@data-testid='review-card-parent']", timeout=10)
            
            review_elements = browser.find_elements("xpath://*[@data-testid='review-card-parent']//div[contains(@class,'ipc-html-content-inner-div')]")
            for i, element in enumerate(review_elements):
                if i >= 5:
                    break
                try:
                    text = element.text.strip()
                    if text:
                        reviews_list.append(f"Review {i+1}: {text}")
                except:
                    continue
        except:
            pass
    
    # Format reviews as single string like Robot Framework
    featured_reviews = "\n---\n".join(reviews_list) if reviews_list else None
    
    return imdb_rating, popularity, metascore, ", ".join(genres) if genres else None, featured_reviews, user_reviews_count



def _insert_row(conn: sqlite3.Connection, title: str, rating: str | None, popularity: str | None,
                metascore: str | None, user_reviews: str | None, genre: str | None, reviews: list, status: str):
    # Convert reviews list back to single string
    featured_reviews = "\n---\n".join([r for r in reviews if r]) if any(reviews) else None
    
    conn.execute(
        """
        INSERT INTO movies (
            title, rating, popularity, metascore, genre, featured_reviews, user_reviews, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (title, rating, popularity, metascore, genre, featured_reviews, user_reviews, status),
    )

def _email_results(db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT * FROM movies")
    rows = cur.fetchall()
    headers = [d[0] for d in cur.description]

    csv_path = os.path.join("output", "movies.csv")
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write(",".join(headers) + "\n")
        for row in rows:
            def esc(val):
                if val is None:
                    return ""
                s = str(val).replace('"', '""')
                if "," in s or "\n" in s:
                    return f'"{s}"'
                return s
            f.write(",".join(esc(v) for v in row) + "\n")

    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASSWORD")
    email_to = os.getenv("EMAIL_TO")
    email_from = os.getenv("EMAIL_FROM", smtp_user or "")
    if not (smtp_host and smtp_user and smtp_pass and email_to and email_from):
        print("Email not sent: missing SMTP credentials or recipient")
        return

    mail = ImapSmtp()
    mail.authorize(smtp_server=smtp_host, smtp_port=smtp_port, account=smtp_user, password=smtp_pass)
    mail.send_message(
        sender=email_from,
        recipients=email_to,
        subject="Movie scrape results",
        body="Find attached the latest movie scrape results.",
        attachments=[csv_path, db_path],
    )

# ------------------- RUN -------------------
if __name__ == "__main__":
    scrape_movies_and_email()
