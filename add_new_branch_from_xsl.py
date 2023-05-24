import requests
import json
import xlrd

# Umbrella API key
management_client_id = "----YOUR MANAGEMENT CLIENT ID HERE----"
management_client_secret = "----YOUR MANAGEMENT CLIENT SECRET HERE----"
network_devices_client_id = "----YOUR NETWORK DEVICE CLIENT ID HERE----"
network_devices_client_secret = "----YOUR NETWORK DEVICE CLIENT SECRET HERE----"

# Umbrella Org ID
umbrella_orgid = '----YOUR UMBRELLA ORG ID HERE----'

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

# Add new site
def add_site(site_name):
    url = "https://management.api.umbrella.com/v1/organizations/{}/sites".format(umbrella_orgid)
    payload = {}
    payload['name'] = site_name
    response = requests.request('GET', url, headers=headers, auth= (management_client_id, management_client_secret))
    if response.status_code == 200:
        if '"name":"{}"'.format(site_name) in response.text :
            for site in json.loads(response.text):
                if site['name'] == site_name:
                    siteId = site['siteId']
                    #print ('Site - {} - already exists; siteId = {}'.format(site_name, siteId))
        else:
            response = requests.request('POST', url, headers=headers, json = payload, auth= (management_client_id, management_client_secret))
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
def add_internal_network(network_name, ipAddress, prefixLength, siteId):
    url = "https://management.api.umbrella.com/v1/organizations/{}/internalnetworks".format(umbrella_orgid)
    payload = {}
    payload['name'] = network_name
    payload['ipAddress'] = ipAddress
    payload['prefixLength'] = prefixLength
    payload['siteId'] = siteId
    response = requests.request('GET', url, headers=headers, auth=(management_client_id, management_client_secret))
    if response.status_code == 200:
        if '"name":"{}"'.format(network_name) in response.text :
            for network in json.loads(response.text):
                if network['name'] == network_name:
                    originId = network['originId']
                    #print ('Internal Network - {} - already exists; originId = {}'.format(network_name, originId))
        else:
            response = requests.request('POST', url, headers=headers, json = payload, auth= (management_client_id, management_client_secret))
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
def add_identity_to_policy(originId, policy_name):
    url = "https://management.api.umbrella.com/v1/organizations/{}/policies".format(umbrella_orgid)
    payload = None
    response = requests.request('GET', url, headers=headers, data=payload,auth=(network_devices_client_id, network_devices_client_secret))
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
    url = "https://management.api.umbrella.com/v1/organizations/{}/policies/{}/identities/{}".format(umbrella_orgid,policyId,originId)
    #time.sleep(15)
    response = requests.request('PUT', url, headers=headers, data=payload,auth=(network_devices_client_id, network_devices_client_secret))
    return response.status_code

# Main def to add new store
def add_store(site_name, network_name, ipAddress, prefixLength, policy_name):
    siteId = add_site(site_name)
    originId = add_internal_network(network_name, ipAddress, prefixLength, siteId)
    response = add_identity_to_policy(originId, policy_name)
    if response == 200:
        message = 'Identity {} was added to {} policy'.format(network_name,policy_name)
    else:
        message = 'Error : not able to add identity {} to policy {}'.format(network_name, policy_name)
    return message

if __name__ == '__main__':
    for row in data:
        site_name = row['site_name']
        policy_name = row['policy_name']
        network_name = row['network_name']
        ipAddress = row['ipAddress']
        prefixLength = row['prefixLength']
        message = add_store(site_name, network_name, ipAddress, prefixLength, policy_name)
        print (message)
