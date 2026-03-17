# News Headlines Scraper

A Python script that automatically fetches recent news headlines from multiple sources and saves them in structured formats.

## How It Works

### 1. **Data Collection**
- **RSS Feed Parsing**: Fetches data from RSS feeds using HTTP requests
- **Multiple Sources**: Collects headlines from TechCrunch and BBC News
- **Real-time Data**: Gets the 4 most recent headlines from each source

## Technical Stack
- Python 3.x

- Requests: HTTP library for API calls

- Pandas: Data manipulation and CSV/Excel export

- ElementTree: XML parsing for RSS feeds

- OpenPyXL: Excel file formatting and styling

## Running the project 
```
pip install -r requirements.txt
python news_scraper.py
```
