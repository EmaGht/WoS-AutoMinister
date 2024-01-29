import requests;
import tempfile;
import time;
from selenium import webdriver

# OneDrive Login Informations
clientId = "bc442fda-48cf-4209-b492-d99b36343e87"
scope = "Files.Read"
redirectUrl = "https://login.microsoftonline.com/common/oauth2/nativeclient"

# Step 1: OneDrive download Excel file for Minister schedulation
URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=' + clientId + '&scope=' + scope + '&response_type=code&redirect_uri=' + redirectUrl

# Step 1.1: Request for login
browser = webdriver.Firefox()
browser.get(URL)
code = ""
while code == "":
    text = browser.current_url
    if text.find('?code') != -1:
        code = text[(text.find('?code') + len('?code') + 1) :]
    time.sleep(1)
browser.quit()

# Step 1.2: Authentication
params = { 'scope': scope, 'code': code, 'client_id': clientId, 'redirect_uri': redirectUrl, 'grant_type': 'authorization_code'}
response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=params)
response.raise_for_status()
header = {'Authorization': 'Bearer ' + response.json().get('access_token')}

# Step 1.3: Lookup for 'Minister_role_scheduler' file between the files shared with the current account
response = requests.get('https://graph.microsoft.com/v1.0/me/drive/sharedWithMe', headers=header)
response.raise_for_status()
file = [x for x in response.json()['value'] if x['name'].startswith("Minister_role_scheduler")][0]

response = requests.get('https://graph.microsoft.com/v1.0/drives/'+ file['parentReference']['driveId'] + '/items/' + file['id'] + '/content', headers=header)
with tempfile.TemporaryFile() as tmpFile:
    tmpFile.write(response.content)

    # Step 2: Open file with Excel reader