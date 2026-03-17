import requests
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
import time

def parse_date(date_string):
    """Parse and format date for Excel compatibility"""
    if not date_string or date_string == "No date":
        return "Unknown"

    # Common RSS date formats
    date_formats = [
        '%a, %d %b %Y %H:%M:%S %z',  # RFC 822 with timezone
        '%a, %d %b %Y %H:%M:%S %Z',  # RFC 822 with timezone name
        '%a, %d %b %Y %H:%M:%S',     # RFC 822 without timezone
        '%Y-%m-%dT%H:%M:%S%z',       # ISO 8601 format
        '%Y-%m-%d %H:%M:%S'          # Simple format
    ]

    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(date_string, fmt)
            # Format for Excel compatibility (YYYY-MM-DD HH:MM:SS)
            return parsed_date.strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            continue

    # If all parsing fails, return the original string
    return date_string

def fetch_techcrunch_headlines():
    """Fetch recent headlines from TechCrunch RSS feed"""
    try:
        url = "https://techcrunch.com/feed/"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        root = ET.fromstring(response.content)
        headlines = []

        # Parse RSS items
        for item in root.findall('.//item')[:4]:
            title = item.find('title').text if item.find('title') is not None else "No title"
            link = item.find('link').text if item.find('link') is not None else "No link"
            pub_date = item.find('pubDate').text if item.find('pubDate') is not None else "No date"

            # Parse and format date
            formatted_date = parse_date(pub_date)

            headlines.append({
                'Source': 'TechCrunch',
                'Title': title,
                'Link': link,
                'Published Date': formatted_date
            })

        return headlines

    except Exception as e:
        print(f"Error fetching TechCrunch headlines: {e}")
        return []

def fetch_bbc_news_headlines():
    """Fetch recent headlines from BBC News RSS feed"""
    try:
        url = "http://feeds.bbci.co.uk/news/rss.xml"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        root = ET.fromstring(response.content)
        headlines = []

        # Parse RSS items
        for item in root.findall('.//item')[:4]:
            title = item.find('title').text if item.find('title') is not None else "No title"
            link = item.find('link').text if item.find('link') is not None else "No link"
            pub_date = item.find('pubDate').text if item.find('pubDate') is not None else "No date"

            # Parse and format date
            formatted_date = parse_date(pub_date)

            headlines.append({
                'Source': 'BBC News',
                'Title': title,
                'Link': link,
                'Published Date': formatted_date
            })

        return headlines

    except Exception as e:
        print(f"Error fetching BBC News headlines: {e}")
        return []

def save_to_csv(headlines, filename='news_headlines.csv'):
    """Save headlines to CSV file with proper encoding and formatting"""
    df = pd.DataFrame(headlines)

    # Ensure proper CSV formatting for Excel
    df.to_csv(filename, index=False, encoding='utf-8-sig')  # utf-8-sig for Excel compatibility

    print(f"Headlines saved to {filename}")
    print(f"Total headlines fetched: {len(headlines)}")
    return df

def save_to_excel_with_formatting(headlines, filename='news_headlines.xlsx'):
    """Save to Excel with proper column formatting"""
    try:
        df = pd.DataFrame(headlines)

        # Create Excel writer
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Headlines')

            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Headlines']

            # Set column widths for better visibility
            worksheet.column_dimensions['A'].width = 15  # Source
            worksheet.column_dimensions['B'].width = 60  # Title
            worksheet.column_dimensions['C'].width = 40  # Link
            worksheet.column_dimensions['D'].width = 20  # Published Date

            # Format the date column
            for row in range(2, len(df) + 2):  # Start from row 2 (skip header)
                worksheet[f'D{row}'].number_format = 'YYYY-MM-DD HH:MM:SS'

        print(f"Headlines saved to Excel file: {filename}")
        print("Column widths optimized for better visibility")
        return df
    except ImportError:
        print("Install 'openpyxl' for Excel support: pip install openpyxl")
        return None

def print_formatted_summary(headlines):
    """Print top 2 headlines in formatted summary"""
    print("\n" + "="*80)
    print("TOP 2 HEADLINES SUMMARY")
    print("="*80)

    # Get top 2 headlines (first 2 from the combined list)
    top_headlines = headlines[:2]

    for i, headline in enumerate(top_headlines, 1):
        print(f"\n{i}. {headline['Source']}")
        print(f"   Title: {headline['Title']}")
        print(f"   Link: {headline['Link']}")
        print(f"   Published: {headline['Published Date']}")
        print("-" * 80)

def main():
    print("Starting News Headlines Scraper...")
    print("Fetching headlines from TechCrunch and BBC News...")

    # Fetch headlines from both sources
    techcrunch_headlines = fetch_techcrunch_headlines()
    bbc_headlines = fetch_bbc_news_headlines()

    # Combine all headlines
    all_headlines = techcrunch_headlines + bbc_headlines

    if all_headlines:
        # Save to CSV (primary)
        df_csv = save_to_csv(all_headlines)

        # Save to Excel with proper formatting
        df_excel = save_to_excel_with_formatting(all_headlines)

        # Print summary
        print_formatted_summary(all_headlines)

        # Display the dataframe
        print("\n📊 All Headlines:")
        print(df_csv.to_string(index=False, max_colwidth=50))


    else:
        print(" No headlines were fetched. Please check your internet connection.")

if __name__ == "__main__":
    main()
