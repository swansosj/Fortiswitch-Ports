import openpyxl
from requests import request
from getpass import getpass
from pprint import pprint
from urllib3 import disable_warnings

baseUrl = "https://10.250.255.254/api/v2/cmdb/"

token = getpass("Please paste your API Token: ")

def getManagedSwitches(token):
    print('Gathering Managed Switch Info')
    url = baseUrl + "/switch-controller/managed-switch"
    payload = {}
    headers = {
        'Authorization': 'Bearer ' + token
    }
    disable_warnings()
    response = request("GET", url, headers=headers, data=payload, verify=False)
    responseJson = response.json()
    result = responseJson['results']

    switchIds = []
    for switch in result:
        switchID = switch['switch-id']
        switchName = switch['name']
        switchInfo = {}
        switchInfo['id'] = switchID
        switchInfo['name'] = switchName
        switchIds.append(switchInfo)
    return switchIds

def getPortVlans(token, switchId, switchName):
    print('Getting Interface Config for ' + switchName)
    url = baseUrl + "/switch-controller/managed-switch/" + switchId
    payload = {}
    headers = {
        'Authorization': 'Bearer ' + token
    }
    disable_warnings()
    response = request("GET", url, headers=headers, data=payload, verify=False)
    responseJson = response.json()
    result = responseJson['results'][0]
    ports = result['ports']

    return ports

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Switch Ports'

sheet['A1'] = 'Switch'
sheet['B1'] = 'Serial Number'
sheet['C1'] = 'Interface'
sheet['D1'] = 'Allowed VLANs'

switchIds = getManagedSwitches(token)

switchCell = 2
serialCell = 2
portCell = 2
vlanCell = 2


for switch in switchIds:
    ports = getPortVlans(token, switch['id'], switch['name'])
    for port in ports:
        switchIDcell = 'A' + str(switchCell)
        serialIDCell = 'B' + str(serialCell)
        portIDCell = 'C' + str(portCell)
        vlanIDCell = 'D' + str(vlanCell)
        portName = port['port-name']
        portAllowedVlans = port['allowed-vlans']
        allowedVlans = ""
        for aVlans in portAllowedVlans:
            allowedVlans = allowedVlans + aVlans['vlan-name'] + ','
        sheet[switchIDcell] = switch['name']
        switchCell = switchCell + 1
        sheet[serialIDCell] = switch['id']
        serailCell = serialCell + 1
        sheet[portIDCell] = portName
        portCell = portCell + 1
        sheet[vlanIDCell] = allowedVlans
        vlanCell = vlanCell + 1

wb.save('FortiSwitch-Interfaces.xlsx')


