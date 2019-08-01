"""Swapify.

Example of how to get credentials: https://gspread.readthedocs.io/en/latest/oauth2.html

TODO:
    * Parameterize Sheet Name
    * ALso allow CSV/Excel file as input
"""
import json
import random
import smtplib

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


def list_derangement(ls):
    """Perform a derangement permutation on a list."""
    ls_random = ls[:]
    while True:
        random.shuffle(ls_random)
        for a, b in zip(ls, ls_random):
            if a == b:
                break
        else:
            return ls_random


def send_email():
    pass


if __name__ == '__main__':

    drive_creds = ServiceAccountCredentials.from_json_keyfile_name('secrets/drive_creds.json', scope)

    gc = gspread.authorize(drive_creds)
    wks = gc.open("Swapify").sheet1

    names = wks.col_values(1)[1:]
    emails = wks.col_values(2)[1:]
    playlists = wks.col_values(3)[1:]

    assert len(playlists) > 1

    playlists_random = list_derangement(playlists)

    with open('secrets/email_creds.json', 'r') as f:
        email_creds = json.load(f)

    server = smtplib.SMTP(host=email_creds['host'], port=email_creds['port'])
    server.starttls()
    server.login(email_creds['email'], email_creds['password'])
    msg = "Test."
    server.sendmail(from_addr=email_creds['email'], to_addrs="ehenry@spins.com", msg=msg)

    # think about how to archive old data
