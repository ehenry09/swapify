"""Swapify

Example of how to get credentials: https://gspread.readthedocs.io/en/latest/oauth2.html

TODO:
    * Potentially filter cells before reading columns into memory
    * Allow local CSV/Excel file as input
        * Parameterize Sheet Name
    * Create archive of old data

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


def build_df_from_gspread(creds, sheet_name="Swapify"):
    """Build a dataframe from Google Sheets API."""
    gc = gspread.authorize(creds)
    spreadsheet = gc.open(sheet_name)
    worksheet = spreadsheet.sheet1

    dtimes = worksheet.col_values(1)[1:]
    names = worksheet.col_values(2)[1:]
    emails = worksheet.col_values(3)[1:]
    playlists = worksheet.col_values(4)[1:]
    col_names = ['dtime', 'name', 'email', 'playlist']
    return pd.DataFrame(zip(dtimes, names, emails, playlists), columns=col_names, dtype=str)


def keep_only_spotify_urls(df, url_prefix=url_prefix):
    """Filter out rows that do not have a correctly formatted URL."""
    return df[df['playlist'].str.startswith(url_prefix)]


def get_most_recent_playlists(df):
    """Filter out rows to exclude."""
    df['date'] = pd.to_datetime(df['dtime']).dt.strftime('%Y-%m-%d')
    week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    df = df[df['date'] > week_ago]
    # handle cases when people submit multiple playlists
    max_mask = df.groupby(['email'])['dtime'].transform(max) == df['dtime']
    return df[max_mask].reset_index(drop=True)


def list_derangement(ls):
    """Perform derangement permutation on a list."""
    assert len(ls) > 1
    ls_random = ls[:]
    while True:
        random.shuffle(ls_random)
        for a, b in zip(ls, ls_random):
            if a == b:
                break
        else:
            return ls_random


def shuffle_playlists(df, playlist_col):
    """Shuffle playlist and append to pd.DataFrame."""
    playlists = df[playlist_col].to_list()
    playlists_random = pd.Series(list_derangement(playlists), name='playlist_to_email')
    df = pd.concat([df, playlists_random], axis=1)
    playlist_from = dict(zip(df['playlist'], df['name']))
    df["playlist_from"] = df['playlist_to_email'].map(playlist_from)
    return df


def send_email(to_addr: str, from_addr: str, recipient: str, curated_by: str, playlist: str):
    """Send the weekly Swapify email."""
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

    df = build_df_from_gspread(creds=drive_creds)
    df = keep_only_spotify_urls(df)
    df = get_most_recent_playlists(df)
    assert len(df) > 1
    df = shuffle_playlists(df, 'playlist')

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
