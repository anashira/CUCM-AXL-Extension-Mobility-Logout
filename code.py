import requests
import base64
import openpyxl

# Define the AXL API URL
url = "https://hq-cucm-pub.ciscouc.com/axl/"

# Define username and password
username = "axlapiuser"
password = "axlapiuser"

# Define request headers, including Basic Authentication
headers = {
    'Content-Type': 'text/xml',
    'Authorization': f'Basic {base64.b64encode(f"{username}:{password}".encode()).decode()}'
}

def device_logout(device_name):
    # Define the SOAP request payload
    payload = f"""
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/12.5">
       <soapenv:Header/>
       <soapenv:Body>
          <ns:doDeviceLogout sequence="">
             <deviceName>{device_name}</deviceName>
          </ns:doDeviceLogout>
       </soapenv:Body>
    </soapenv:Envelope>
    """

    try:
        # Send a POST request to the AXL API
        response = requests.post(url, headers=headers, data=payload, verify=False)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            print(f"Request for device '{device_name}' was successful. Response:")
            print(response.text)
        else:
            print(f"Request for device '{device_name}' failed with status code: {response.status_code}")
            print(response.text)
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while sending the request: {e}")

if __name__ == "__main__":
    # Load the Excel file
    workbook = openpyxl.load_workbook('data.xlsx')
    sheet = workbook.active

    # Iterate through rows in the Excel file to get device names
    for row in sheet.iter_rows(values_only=True):
        device_name = row[0]
        
        # Skip empty rows
        if device_name is None:
            continue
        
        device_logout(device_name)

    workbook.close()  # Close the Excel file after usage
