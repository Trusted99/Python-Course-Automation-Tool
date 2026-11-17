
import csv
from pathlib import Path
from datetime import timedelta

import yt_dlp


def format_duration(seconds):
    """Convert video duration in seconds to HH:MM:SS or MM:SS."""
    if seconds is None:
        return "Unavailable"

    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    if hours:
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    return f"{minutes:02}:{seconds:02}"


def extract_playlist(playlist_url):
    """Extract playlist videos and preserve links when possible."""

    options = {
        "extract_flat": True,
        "skip_download": True,
        "ignoreerrors": True,
        "quiet": True,
        "no_warnings": True,
        "noplaylist": False,
    }

    videos = []

    print("\nReading YouTube playlist. Please wait...\n")

    with yt_dlp.YoutubeDL(options) as ydl:
        playlist = ydl.extract_info(
            playlist_url,
            download=False
        )

        if not playlist:
            raise RuntimeError(
                "Could not read the playlist. Check the URL."
            )

        entries = playlist.get("entries") or []
        total = len(entries)

        for index, entry in enumerate(entries, start=1):
            if not entry:
                print(
                    f"Entry {index}/{total}: "
                    "No video ID or metadata available; skipped."
                )
                continue

            video_id = entry.get("id")
            title = entry.get("title") or "Unavailable video"
            duration = entry.get("duration")

            # Recover the ID from the entry URL if necessary.
            entry_url = entry.get("url") or entry.get("webpage_url")

            if not video_id and entry_url:
                if "youtube.com/watch?v=" in entry_url:
                    video_id = entry_url.split("v=", 1)[1].split("&", 1)[0]
                elif "youtu.be/" in entry_url:
                    video_id = entry_url.split("youtu.be/", 1)[1].split("?")[0]
                elif entry_url.isalnum() and len(entry_url) == 11:
                    video_id = entry_url

            # Build a direct link whenever an ID is available.
            video_link = (
                f"https://www.youtube.com/watch?v={video_id}"
                if video_id
                else entry.get("webpage_url")
            )

            # Attempt to retrieve complete metadata.
            if video_id and (
                duration is None
                or title in ("Unavailable video", "Private video")
            ):
                try:
                    with yt_dlp.YoutubeDL({
                        "skip_download": True,
                        "quiet": True,
                        "no_warnings": True,
                        "ignoreerrors": True,
                    }) as detail_ydl:
                        details = detail_ydl.extract_info(
                            video_link,
                            download=False
                        )

                    if details:
                        title = details.get("title") or title

                        if details.get("duration") is not None:
                            duration = details["duration"]

                        video_link = (
                            details.get("webpage_url") or video_link
                        )

                except Exception:
                    # Keep the direct link even if metadata is inaccessible.
                    pass

            # Preserve entries when a video link is known.
            if video_link:
                videos.append({
                    "Video Title": title,
                    "Duration": format_duration(duration),
                    "Shareable Video Link": video_link,
                })

                print(
                    f"Processed {index}/{total}: "
                    f"{title} | {video_link}"
                )
            else:
                print(
                    f"Entry {index}/{total}: "
                    "Could not recover a video link."
                )

    return videos, playlist.get("title", "youtube_playlist")


def save_results(videos, playlist_title):
    """Save the extracted data as Excel and CSV files."""

    if not videos:
        print("\nNo video information was extracted.")
        return

    # Create a safe filename from the playlist title.
    safe_title = "".join(
        c if c.isalnum() or c in (" ", "-", "_") else "_"
        for c in playlist_title
    ).strip()

    safe_title = safe_title[:100] or "youtube_playlist"

    output_dir = Path("playlist_exports")
    output_dir.mkdir(exist_ok=True)

    csv_path = output_dir / f"{safe_title}.csv"
    excel_path = output_dir / f"{safe_title}.xlsx"

    columns = [
        "Video Title",
        "Duration",
        "Shareable Video Link",
    ]

    # Export CSV.
    with open(
        csv_path, "w", newline="", encoding="utf-8-sig"
    ) as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(videos)

    # Export Excel.
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Playlist Videos"

    sheet.append(columns)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    for video in videos:
        sheet.append([video[column] for column in columns])

    # Make the links clickable in Excel.
    for row in range(2, sheet.max_row + 1):
        link_cell = sheet.cell(row=row, column=3)
        link_cell.hyperlink = link_cell.value
        link_cell.style = "Hyperlink"

    widths = [55, 16, 65]
    for column_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            get_column_letter(column_index)
        ].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(excel_path)

    print("\nExtraction completed successfully!")
    print(f"Total videos extracted: {len(videos)}")
    print(f"Excel file: {excel_path.resolve()}")
    print(f"CSV file:   {csv_path.resolve()}")


def main():
    print("=" * 55)
    print("       YOUTUBE PLAYLIST VIDEO EXTRACTOR")
    print("=" * 55)

    playlist_url = input(
        "\nPaste your YouTube playlist URL: "
    ).strip()

    if not playlist_url:
        print("No URL provided. Exiting.")
        return

    try:
        videos, playlist_title = extract_playlist(playlist_url)
        save_results(videos, playlist_title)

    except Exception as error:
        print(f"\nExtraction failed: {error}")
        print(
            "Check your internet connection, playlist URL, "
            "and yt-dlp version. Some playlists require "
            "authentication or may restrict access."
        )


if __name__ == "__main__":
    main()
