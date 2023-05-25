from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
import requests
import json
import xlrd

# Export/Set API key
client_id = "---YOUR UMBRELLA CLIENT ID HERE---"
client_secret = "---YOUR UMBRELLA API secrect HERE---"

# Open the Workbook
workbook = xlrd.open_workbook('branch.xls')

# Open the worksheet
worksheet = workbook.sheet_by_index(0)

no_of_rows = worksheet.nrows

# Iterate the rows and columns
data = []
while no_of_rows > 0:
    for row in range(no_of_rows):
        line = {}
        if row != 0:
            for columns in range(1):
                line['site_name'] = worksheet.cell_value(row, columns)
            for columns in range(2):
                line['policy_name'] = worksheet.cell_value(row, columns)
            for columns in range(3):
                line['network_name'] = worksheet.cell_value(row, columns)
            for columns in range(4):
                line['ipAddress'] = worksheet.cell_value(row, columns)
            for columns in range(5):
                line['prefixLength'] = worksheet.cell_value(row, columns)
            data.append(line)
        no_of_rows = no_of_rows - 1

headers = {
    "Content-Type": "application/json",
    "Accept": "application/json"
}

class UmbrellaAPI:
    def __init__(self, url, ident, secret):
        self.url = url
        self.ident = ident
        self.secret = secret
        self.token = None

    def GetToken(self):
        auth = HTTPBasicAuth(self.ident, self.secret)
        client = BackendApplicationClient(client_id=self.ident)
        oauth = OAuth2Session(client=client)
        self.token = oauth.fetch_token(token_url=self.url, auth=auth)
        return self.token


# Add new site
def add_site(token, site_name):
    url = "https://api.umbrella.com/deployments/v2/sites"
    payload = {}
    payload['name'] = site_name
    headers['Authorization'] = 'Bearer '+token
    response = requests.request('GET', url, headers=headers)
    if response.status_code == 200:
        if '"name":"{}"'.format(site_name) in response.text :
            for site in json.loads(response.text):
                if site['name'] == site_name:
                    siteId = site['siteId']
                    #print ('Site - {} - already exists; siteId = {}'.format(site_name, siteId))
        else:
            response = requests.request('POST', url, headers=headers, json=payload)
            if response.status_code == 200:
                siteId = json.loads(response.text)['siteId']
                    #print ('New site - {} - was created with siteId :{}'.format(site_name, siteId))
            else:
                print ('Error - Not able to add new site {}'.format(site_name))
                exit()

    else:
        print('Error - Not able to get site liste from Umbrella API'.format(site_name))
        exit()

    return siteId

# Add new internal network
def add_internal_network(token, network_name, ipAddress, prefixLength, siteId):
    url = "https://api.umbrella.com/deployments/v2/internalnetworks"
    payload = {}
    payload['name'] = network_name
    payload['ipAddress'] = ipAddress
    payload['prefixLength'] = prefixLength
    payload['siteId'] = siteId
    headers['Authorization'] = 'Bearer '+token
    response = requests.request('GET', url, headers=headers)
    if response.status_code == 200:
        if '"name":"{}"'.format(network_name) in response.text :
            for network in json.loads(response.text):
                if network['name'] == network_name:
                    originId = network['originId']
                    #print ('Internal Network - {} - already exists; originId = {}'.format(network_name, originId))
        else:
            response = requests.request('POST', url, headers=headers, json = payload)
            if response.status_code == 200:
                originId = json.loads(response.text)['originId']
                #print ('New Internal Network - {} - was created with originId :{}'.format(network_name, originId))
            else:
                print ('Error - Not able to add new internal network {}'.format(network_name))
                exit()

    else:
        print('Error - Not able to get internal networks list from Umbrella API'.format(site_name))
        exit()

    return originId

# Add identity to policy
def add_identity_to_policy(token, originId, policy_name):
    #url = "https://management.api.umbrella.com/v1/organizations/{}/policies".format(umbrella_orgid)
    url = "https://api.umbrella.com/deployments/v2/policies"
    payload = None
    headers['Authorization'] = 'Bearer '+token
    response = requests.request('GET', url, headers=headers, data=payload)
    if response.status_code == 200:
        if '"name":"{}"'.format(policy_name) in response.text:
            for policy in json.loads(response.text):
                if policy['name'] == policy_name:
                    policyId = policy['policyId']
                    #print ('Policy Name - {} - policyId = {}'.format(policy_name, policyId))
        else:
            print ('Policy - {} - does not exist'.format(policy_name))
            exit()
    else:
        print('Error - Not able to get policy list from Umbrella API')
        exit()
    url = "https://api.umbrella.com/deployments/v2/policies/{}/identities/{}".format(policyId, originId)
    #time.sleep(15)
    response = requests.request('PUT', url, headers=headers, data=payload)
    return response.status_code

# Main def to add new store
def add_store(token, site_name, network_name, ipAddress, prefixLength, policy_name):
    siteId = add_site(token, site_name)
    originId = add_internal_network(token, network_name, ipAddress, prefixLength, siteId)
    response = add_identity_to_policy(token, originId, policy_name)
    if response == 200:
        message = 'Identity {} was added to {} policy'.format(network_name,policy_name)
    else:
        message = 'Error : not able to add identity {} to policy {}'.format(network_name, policy_name)
    return message

if __name__ == '__main__':
    token_url = 'https://api.umbrella.com/auth/v2/token'
    api = UmbrellaAPI(token_url, client_id, client_secret)
    token = api.GetToken()["access_token"]
    for row in data:
        site_name = row['site_name']
        policy_name = row['policy_name']
        network_name = row['network_name']
        ipAddress = row['ipAddress']
        prefixLength = row['prefixLength']
        message = add_store(token, site_name, network_name, ipAddress, prefixLength, policy_name)
        print (message)
