"""Swapify.

Example of how to get credentials: https://gspread.readthedocs.io/en/latest/oauth2.html

TODO:
    * Parameterize Sheet Name
    * ALso allow CSV/Excel file as input
"""
import random

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


def list_derangement(ls):
    """Perform a derangement permutation on a list.

    Args:
        ls (list): List to undergo derangement.

    Returns:
        A randomized list with no item maintaining its original position.

    """
    ls_random = ls[:]
    while True:
        random.shuffle(ls_random)
        for a, b in zip(ls, ls_random):
            if a == b:
                break
        else:
            return ls_random


if __name__ == '__main__':

    credentials = ServiceAccountCredentials.from_json_keyfile_name('secrets/creds.json', scope)

    gc = gspread.authorize(credentials)
    wks = gc.open("Swapify").sheet1

    names = wks.col_values(1)[1:]
    emails = wks.col_values(2)[1:]
    playlists = wks.col_values(3)[1:]

    playlists_random = list_derangement(playlists)

    ########################
    # NEED TO UPDATE BELOW #
    ########################
    
    email = dfrand["What's your e-mail?"].values
    url = dfrand["What's your Spotify playlist URL?"].values
    dictionary = dict(zip(email, url))  # make dictionary out of randomized lists

    # Function to write and send e-mails
    def send_notification():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = 'Playlist Day!'
        mail.body = ("It's playlist day! Here's a playlist:\n\n" +
                     url + "\n \n"
                     "To reach your playlist, enter the url into the Spotify search bar and search it. \n \n"
                     "If you got your own playlist back, let me know and he'll get you a different one. \n \n"
                     "Thanks for participating!")
        mail.send

    # Send e-mail
    for email, url in dictionary.items():
        send_notification()
