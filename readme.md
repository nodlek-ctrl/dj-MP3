# YouTube Lyric Video Downloader

This program searches for lyric videos of songs on YouTube and downloads them as MP3s.

## Requirements

- Python 3.x
- `openpyxl` Python module (`pip install openpyxl`)
- `youtube-search-python` Python module (`pip install youtube-search-python`)
- `yt-dlp` Python module (`pip install yt-dlp`)

## Usage

1. Create a new spreadsheet called `songs.xlsx` in the same directory as the program.
2. In the spreadsheet, create two columns: `Song` and `Artist`. The first row of the spreadsheet should contain headers for the columns.
3. Fill in the `Song` and `Artist` columns with the names of the songs and artists you want to search for.
4. Run the program using the command `python main.py`.
5. The program will search for a matching lyric video for each song and artist combination in the spreadsheet. If a matching video is found, the program will retrieve the video ID and store it in the sixth column of the same row.
6. If the user chooses to download the video, the program uses the video ID to download the video as an MP3.

## Spreadsheet Layout

The `songs.xlsx` spreadsheet has two columns: `Song` and `Artist`. The first row of the spreadsheet contains headers for the columns.

Here's an example of what the spreadsheet might look like:

| Song             | Artist           |
|------------------|------------------|
| Don't Start Now  | Dua Lipa         |
| Blinding Lights  | The Weeknd       |
| Levitating       | Dua Lipa feat.   |

Each row of the spreadsheet represents a different song. The program uses the `Song` and `Artist` columns to search for a matching lyric video on YouTube. If a matching video is found, the program retrieves the video ID and stores it in the sixth column of the same row. If the user chooses to download the video, the program uses the video ID to download the video as an MP3.
