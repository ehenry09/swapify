"""Swapify

Example of how to get credentials: https://gspread.readthedocs.io/en/latest/oauth2.html

TODO:
    * Potentially filter cells before reading into memory
    * Allow local CSV/Excel file as input
    * Parameterize Sheet Name
    * Create archive of old data
    * Check out Spotify URIs
    * Write Reminder Email
"""
import json
import random
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from smtplib import SMTPRecipientsRefused

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

url_prefix = "https://open.spotify.com/playlist/"
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


def list_derangement(ls):
    """Perform derangement permutation on a list."""
    ls_random = ls[:]
    while True:
        random.shuffle(ls_random)
        for a, b in zip(ls, ls_random):
            if a == b:
                break
        else:
            return ls_random


def send_email(to_addr: str, from_addr: str, recipient: str, curated_by: str, playlist: str):
    """Send an email."""
    msg = EmailMessage()
    msg['Subject'] = 'Your playlist has arrived!'
    msg['To'] = to_addr
    msg['From'] = from_addr
    body = f"""

    Hey {recipient}!

    Below you'll find your weekly playlist, freshly curated by {curated_by}.

    {playlist}

    Enjoy!

    """
    msg.set_content(body)
    server.send_message(msg)


if __name__ == '__main__':

    drive_creds = ServiceAccountCredentials.from_json_keyfile_name('secrets/drive_creds.json', scope)

    gc = gspread.authorize(drive_creds)
    spreadsheet = gc.open('Swapify')
    worksheet = spreadsheet.sheet1

    dtimes = worksheet.col_values(1)[1:]
    names = worksheet.col_values(2)[1:]
    emails = worksheet.col_values(3)[1:]
    playlists = worksheet.col_values(4)[1:]
    col_names = ['date', 'name', 'email', 'playlist']
    df = pd.DataFrame(zip(dtimes, names, emails, playlists), columns=col_names, dtype=str)

    df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
    week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

    df = df[df['date'] > week_ago]
    # df = df[df['playlist'].str.startswith(url_prefix)]

    assert len(df) > 1

    playlists_random = pd.Series(list_derangement(playlists), name='playlist_to_email')
    df = pd.concat([df, playlists_random], axis=1)

    playlist_from = dict(zip(df['playlist'], df['name']))
    df["playlist_from"] = df['playlist_to_email'].map(playlist_from)

    with open('secrets/email_creds.json', 'r') as f:
        email_creds = json.load(f)

    server = smtplib.SMTP(host=email_creds['host'], port=email_creds['port'])
    server.starttls()
    server.login(email_creds['email'], email_creds['password'])

    for row in df.itertuples(index=False):
        try:
            send_email(to_addr=row.email, from_addr=email_creds['email'], recipient=row.name,
                       curated_by=row.playlist_from, playlist=row.playlist_to_email)
        except SMTPRecipientsRefused:
            pass
