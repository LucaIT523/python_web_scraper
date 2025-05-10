# 

<div align="center">
   <h1>python_web_scraper</h1>
</div>



This code is for TCE order data scraping and Google Sheets integration system.



### . **System Architecture**

#### 1.1 Core Components

```
# Web Automation
from selenium import webdriver
# Data Parsing
from bs4 import BeautifulSoup
# Google API Integration
from googleapiclient import discovery
# Authentication
from oauth2client import file, client, tools
```

- **Browser Automation**: Uses Selenium WebDriver for dynamic web interaction
- **HTML Parsing**: BeautifulSoup for extracting table data
- **Cloud Integration**: Google Sheets API for data synchronization

#### 1.2 Data Mapping Configuration

```
extracting_info = {
    13: 'Nume destinatar',   # Recipient Name
    14: 'Țară destinatar',   # Destination Country
    25: 'Descriere marfă',   # Goods Description
    27: 'Referință expeditor', # Shipper Reference
    21: 'Valoare ramburs'    # COD Value
}
```

- **Column Mapping**: Defines target data columns from TCE order table
- **Multilingual Support**: Contains Romanian language field names

### 2. **Workflow Implementation**

#### 2.1 Authentication Flow

```
pythonCopydef get_credentials():
    # OAuth2 token management
    store = oauth2client.file.Storage('tokens.json')
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        credentials = tools.run_flow(flow, store, flags)
    return credentials
```

- **Token Caching**: Uses local `tokens.json` for credential storage
- **Auto-Refresh**: Handles token expiration automatically

#### 2.2 Data Extraction Process

```
pythonCopydef get_extract_all(driver, page_url):
    # Table parsing logic
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find('table')
    for j in table.find_all('tr')[1:]:
        # Data extraction implementation
        row_id = re.search(r'>(.*?)<', str(th_data[0])).group(1)
        mydata_id.append(row_id)
        # ... (data processing)
```

- **Pagination Handling**: Iterates through all order pages
- **Regex Filtering**: Extracts order IDs using regular expressions

### 3. **Google Sheets Integration**

#### 3.1 Data Synchronization

```
def set_googlesheets_data(service, ID, data):
    # Data append operation
    rangeName = 'A1:F1'
    row_data = {'values': [[ID, data[0], data[1], data[2], data[3], data[4]]]}
    service.spreadsheets().values().append(
        spreadsheetId=SpreadSheetID,
        body=row_data
    ).execute()
```

- **Batch Writing**: Processes data in memory before cloud sync
- **Rate Limiting**: 1-second delay between write operations

#### 3.2 Conflict Resolution

```
pythonCopyIDList = get_googlesheets_IDList(service)
for tce_id in mydata_id:
    if tce_id not in IDList:
        # Write new records only
        set_googlesheets_data(service, tce_id, mydata_row[num])
```

- **Deduplication Check**: Compares local and cloud IDs
- **Delta Update**: Only writes new/changed records

### 4. **Key Features**

#### 4.1 Security Implementation

- **Credential Isolation**: Separate files for TCE (`tce_login_user/pw`) and Google (`creds.json`)
- **Headless Operation**: `driver.minimize_window()` for background execution

#### 4.2 Error Handling

```
try:
    # Main execution block
except HttpError as err:
    print(f'Error ... Google Sheets ({err})')
```

- **API Error Capture**: Handles Google Sheets specific errors
- **Fault Tolerance**: Continues processing after individual record failures

#### 4.3 Performance Optimization

- **Cached Authentication**: Reuses stored OAuth2 tokens
- **Batch Processing**: Processes all pages before cloud sync











### **Contact Us**

For any inquiries or questions, please contact us.

telegram : @topdev1012

email :  skymorning523@gmail.com

Teams :  https://teams.live.com/l/invite/FEA2FDDFSy11sfuegI







