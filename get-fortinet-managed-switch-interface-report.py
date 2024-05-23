import csv
import pandas as pd
from shlex import join
from requests import request
from getpass import getpass
from urllib3 import disable_warnings
 
baseUrl = "https://10.250.255.254/api/v2/cmdb/"
 
disable_warnings()
 
def load(token):
    print('Gathering Managed Switch Info')
 
    url = baseUrl + "/switch-controller/managed-switch"
    headers = {
        'Authorization': 'Bearer ' + token
    }
    response = request("GET", url, headers=headers, verify=False)
    switches = response.json()['results']
    return switches
 
 
def transform(switches):
    """
    Transforms a list of switches into a formatted output.
 
    Args:
        switches (list): A list of dictionaries representing switches.
 
    Returns:
        list: A list of dictionaries containing the transformed data.
 
    Example:
        switches = [
            {
                'switch-id': 1,
                'name': 'Switch 1',
                'ports': [
                    {
                        'port-name': 'Port 1',
                        'allowed-vlans': [
                            {'vlan-name': 'VLAN1'},
                            {'vlan-name': 'VLAN2'}
                        ]
                    },
                    {
                        'port-name': 'Port 2',
                        'allowed-vlans': [
                            {'vlan-name': 'VLAN3'},
                            {'vlan-name': 'VLAN4'}
                        ]
                    }
                ]
            },
            {
                'switch-id': 2,
                'name': 'Switch 2',
                'ports': [
                    {
                        'port-name': 'Port 3',
                        'allowed-vlans': [
                            {'vlan-name': 'VLAN5'},
                            {'vlan-name': 'VLAN6'}
                        ]
                    },
                    {
                        'port-name': 'Port 4',
                        'allowed-vlans': [
                            {'vlan-name': 'VLAN7'},
                            {'vlan-name': 'VLAN8'}
                        ]
                    }
                ]
            }
        ]
 
        transform(switches)
 
    Output:
        [
            {
                'id': 1,
                'name': 'Switch 1',
                'port_name': 'Port 1',
                'port_allowed_vlans': 'VLAN1,VLAN2'
            },
            {
                'id': 1,
                'name': 'Switch 1',
                'port_name': 'Port 2',
                'port_allowed_vlans': 'VLAN3,VLAN4'
            },
            {
                'id': 2,
                'name': 'Switch 2',
                'port_name': 'Port 3',
                'port_allowed_vlans': 'VLAN5,VLAN6'
            },
            {
                'id': 2,
                'name': 'Switch 2',
                'port_name': 'Port 4',
                'port_allowed_vlans': 'VLAN7,VLAN8'
            }
        ]
    """
    output = []
    for s in switches:
        ports = s.get('ports', [])
 
        for port in ports:
            port_allowed_vlans = port.get('allowed-vlans', [])
 
            port_allowed_vlans = [v['vlan-name'] for v in port_allowed_vlans]
 
            output.append(
                {
                    'id': s['switch-id'],
                    'name': s['name'],
                    'port_name': port['port-name'],
                    'port_allowed_vlans': port_allowed_vlans
                }
            )
 
    return output
 
 
def save(output):
    file_path = 'FortiSwitch-Interfaces.csv'
    with open(file_path, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['switch_id', 'switch_name', 'port_name', 'allowed_vlans'])
        for o in output:
            csvwriter.writerow([o['id'], o['name'], o['port_name'], ','.join(o['port_allowed_vlans'])])
    pd.read_csv(file_path)
    df = pd.read_csv(file_path)
    df.to_excel('FortiSwitch-Interfaces.xlsx', index=False)

 
 
def main():
    token = getpass("Please paste your API Token: ")
 
    switches = load(token)
    output = transform(switches)
    save(output)

if __name__ == "__main__":
    main()