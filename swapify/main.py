import csv
import datetime
import glob
import numpy as np
import os
import pandas as pd
import psutil
import random as random
import re
import subprocess
import win32com.client as win32
import zipfile
from collections import Counter
now = datetime.datetime.now()

os.chdir()

# access sheet directly

# load data
# move zipped google forms csv file from downloads to swapify directory
# swapraw = max(glob.glob('C:\\Users\\USER\\Downloads\\FILENAME*.zip'),
#              key = os.path.getctime)
# swapzipped = ('C:\\Users\\USER\\Desktop\\Swapify\\' +
#               re.search('Swapify\_.*\.zip', swapdl).group())
# os.rename(swapraw, swapzipped)

# unzip file
zipped = glob.glob('C:\\Users\\USER\\Desktop\\Swapify\\*.zip',)
zip_ref = zipfile.ZipFile(zipped[0], 'r')
zip_ref.extractall("C:\\Users\\USER\\Desktop\\Swapify")
zip_ref.close()

# rename file

swap = re.search('Swapify\_.*\.csv', zipped[0]).group()
swaprename = 'test' + swap[:-4] + now.strftime('%Y%m%d') + '.csv'
os.rename(swap, swaprename)

# read csv and convert to df

ifile = pd.read_csv(swaprename, usecols=[1, 2])
df = pd.DataFrame(ifile)

# randomize dfs

dfrand = df.apply(np.random.permutation)  # Randomize df rows

# checking for duplicates

check = pd.concat([df, dfrand], join='outer')
dupes = sum(check.duplicated())
while dupes > 0:
    dfrand = df.apply(np.random.permutation)
    check = pd.concat([df, dfrand], join='outer')
    dupes = sum(check.duplicated())


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
