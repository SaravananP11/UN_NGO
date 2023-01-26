import urllib
from _csv import writer

from bs4 import BeautifulSoup
import csv
import requests
import re
import json
import requests
import pandas as pd
import urllib3
from selenium import webdriver
from time import sleep


html = open(r"D:\pythonProject\NGO_html.txt", "r")
r = html.read()
soup = BeautifulSoup(r, 'html5lib') # If this line causes an error, run 'pip install html5lib' or install html5lib
url = []
for a in soup.find_all('a', href=True):
    url.append(a['href'])
with open(r"D:\pythonProject\ngo_url.csv", 'a') as f_object:
    # Pass this file object to csv.writer()
    # and get a writer object
    writer_object = writer(f_object)

    # Pass the list as an argument into
    # the writerow()
    for u in url:
        writer_object.writerow([u])

    # Close the file object
    f_object.close()
