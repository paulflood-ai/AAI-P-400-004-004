import os
import subprocess
import time
from datetime import datetime
import tempfile

import pyautogui
import pyperclip
import gspread
from google.oauth2.service_account import Credentials

def capture_youtube_screenshot_and_paste_to_sheets(youtube_url, sheet_name, cell_to_paste):
    """
    Captures a screenshot of a YouTube video, copies it to the clipboard,
    and pastes it into a specified Google Sheets cell.

    Args:
        youtube_url (str): The URL of the YouTube video.
        sheet_name (str): The name of the Google Sheet.
        cell_to_paste (str): The cell where the screenshot should be pasted (e.g., "A1").
    """
    try:
        # 1. Open YouTube video in Safari (or default browser)
        subprocess.run(["open", "-a", "Safari", youtube_url], check=True) # or change Safari to your default browser

        # 2. Give the browser some time to load the video. Adjust as needed.
        time.sleep(5)  # Adjust the sleep time based on internet speed

        # 3. Take a screenshot of the entire screen.
        screenshot = pyautogui.screenshot()

        # 4. Save the screenshot to a temporary file.
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
            screenshot.save(temp_file.name)
            temp_file_path = temp_file.name

        #5. copy the file to clipboard.
        subprocess.run(["osascript", "-e", f'set the clipboard to (read (POSIX file "{temp_file_path}") as «class PNGf»)'], check=True) # Copies the image to clipboard.

        # 6. Authenticate with Google Sheets API.
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        credentials = Credentials.from_service_account_file(
            "service_account.json", scopes=scopes
        ) # Ensure you have a service_account.json file in the same directory.
        gc = gspread.authorize(credentials)

        # 7. Open the Google Sheet.
        sheet = gc.open(sheet_name).sheet1 # or open the desired sheet by name

        # 8. Paste the screenshot into the specified cell.
        sheet.update_cell(cell_to_paste, "") # Clear the cell first.
        sheet.paste_image(cell_to_paste, temp_file_path)

        #9. delete the temp file.
        os.unlink(temp_file_path)

        print(f"Screenshot pasted into {sheet_name}!{cell_to_paste}")

    except subprocess.CalledProcessError as e:
        print(f"Error opening browser: {e}")
    except FileNotFoundError:
        print("Error: service_account.json not found. Make sure you have the service account credentials.")
    except gspread.exceptions.APIError as e:
        print(f"Google Sheets API error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Example usage:
youtube_url = "https://www.youtube.com/watch?v=FmtJ22fhOvw"  # Replace with your YouTube URL
sheet_name = "My YouTube Screenshots"  # Replace with your Google Sheet name
cell_to_paste = "A1"

capture_youtube_screenshot_and_paste_to_sheets(youtube_url, sheet_name, cell_to_paste)