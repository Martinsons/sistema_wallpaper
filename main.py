import pandas as pd
from states import us_states
from states import us_state_to_abbrev
from countries import country_abbreviations
from boxSizes import box_sizes
from serviceType import service_type_mappings
from statesCA import canadian_provinces
from statesCA import ca_states
from datetime import datetime
# from conditions import conditions
today_date = datetime.today().strftime('%Y-%m-%d')
# Load the Excel file
file_path = 'inbound.xlsx'
df1 = pd.read_excel(file_path)
# test = pd.read_excel(file_path).dtypes

# Create a new DataFrame with specified column names
columns = ['shipmentReference', 'senderContactName', 'senderCompany', 'senderContactNumber', 'senderLine1', 'senderPostcode', 'senderCity', 'senderCountry', 'senderEmail', 'recipientContactName', 'recipientCompany', 'recipientContactNumber', 'recipientLine1', 'recipientLine2', 'recipientState', 'recipientPostcode', 'recipientCity', 'recipientCountry', 'recipientEmail', 'totalShipmentWeight', 'purposeOfShipment', 'serviceType', 'weightUnits', 'numberOfPackages', 'packageType', 'width', 'length', 'height', 'packageWeight', 'carriageValue', 'currencyType', 'itemDescription', 'customsControlled', 'generateInvoice', 'termsOfSale', 'etdEnabled', 'manufacturingCountry', 'commodityQuantity', 'commodityMeasureUnit', 'commodityWeight', 'customsValue', 'commodityType', 'DocumentType', 'DocumentDescription']
df2 = pd.DataFrame(columns=columns)

destination_file = rf'Fedex_batch_{today_date}.xlsx'

# columns_to_copy = df1[['StoreName', 'OrderId', 'CustomerFullName', 'CustomerEmail', 'CustomerPhoneNr', 'DeliveryAddress1', 'DeliveryAddress2', 'DeliveryZip', 'DeliveryCountry', 'DeliveryState', 'DeliveryCity', 'ShippingMethod']]
# Ensure 'BoxSize' column is of string type
df1['BoxSize'] = df1['BoxSize'].astype(str).astype(int)
# Filter out rows where 'BoxSize' has no value
df1 = df1[df1['BoxSize'].notna()]


# Copy only the filtered rows
columns_to_copy = df1[['StoreName', 'OrderId', 'CustomerFullName', 'CustomerEmail', 'CustomerPhoneNr', 'DeliveryAddress1', 'DeliveryAddress2', 'DeliveryZip', 'DeliveryCountry', 'DeliveryState', 'DeliveryCity', 'ShippingMethod']]

# Izvelamies uz kuram kolonam parkopesim datus
# DATU KOPESANA UZ JAUNO FAILU
df2['senderContactName'] = columns_to_copy['StoreName']
df2['senderCompany'] = columns_to_copy['StoreName']
df2['shipmentReference'] = columns_to_copy['OrderId']
df2['recipientContactName'] = columns_to_copy['CustomerFullName']
df2['recipientEmail'] = columns_to_copy['CustomerEmail']
df2['recipientContactNumber'] = columns_to_copy['CustomerPhoneNr']
df2['recipientLine1'] = columns_to_copy['DeliveryAddress1']
df2['recipientLine2'] = columns_to_copy['DeliveryAddress2']
df2['recipientPostcode'] = columns_to_copy['DeliveryZip']
df2['recipientCity'] = columns_to_copy['DeliveryCity']
df2['recipientCountry'] = columns_to_copy['DeliveryCountry']
df2['recipientState'] = columns_to_copy['DeliveryState']
df2['serviceType'] = columns_to_copy['ShippingMethod']

# AIZPILDA TUKSOS LAUKUMUS AR NEPIECIESAMO INFORMACIJU
# Numurs
our_number = ['27071150'] * len(df2)
df2['senderContactNumber'] = our_number
# Adrese
our_adress = ['Brivibas gatve 323'] * len(df2)
df2['senderLine1'] = our_adress
# Postcode
our_postCode = ['1006'] * len(df2)
df2['senderPostcode'] = our_postCode
# Pilseta
our_city = ['Riga'] * len(df2)
df2['senderCity'] = our_city
# Valsts kods
our_countryCode = ['LV'] * len(df2)
df2['senderCountry'] = our_countryCode
# Sutijuma svars un pacinas svars un preces svars
our_shipmentWeight = ['2'] * len(df2)
df2['totalShipmentWeight'] = our_shipmentWeight
df2['packageWeight'] = our_shipmentWeight
df2['commodityWeight'] = our_shipmentWeight
# Sutijuma merkis
our_shipmentPurpose = ['GIFT'] * len(df2)
df2['purposeOfShipment'] = our_shipmentPurpose
# Sutijuma svara mervieniba
our_shipmentKG = ['KGS'] *  len(df2)
df2['weightUnits'] = our_shipmentKG
# Pacinu daudzums un precu daudzums
our_packageCount = ['1'] *  len(df2)
df2['numberOfPackages'] = our_packageCount
df2['commodityQuantity'] = our_packageCount
# Pacinas tips
our_packageType = ['YOUR_PACKAGING'] *  len(df2)
df2['packageType'] = our_packageType
# Augstums
our_packageHeight = ['5'] *  len(df2)
df2['height'] = our_packageHeight
# Valutas veids
our_packageCurrency = ['EUR'] *  len(df2)
df2['currencyType'] = our_packageCurrency
# Sutijuma apraksts
our_packageItem = ['Wallpaper'] *  len(df2)
df2['itemDescription'] = our_packageItem
# Sutijam kontrole un ETD Enable
our_shipmentControlled = ['Y'] *  len(df2)
df2['customsControlled'] = our_shipmentControlled
df2['etdEnabled'] = our_shipmentControlled
# Sutijama invoice
our_generateInvoice = ['PI'] *  len(df2)
df2['generateInvoice'] = our_generateInvoice
# Terms of Sale
our_termsOfSale = ['4'] *  len(df2)
df2['termsOfSale'] = our_termsOfSale
# Razotaja valsts
our_manufacturingCountry = ['LV'] *  len(df2)
df2['manufacturingCountry'] = our_manufacturingCountry
# Preces mervieniba
our_commodityMeasure = ['PCS'] *  len(df2)
df2['commodityMeasureUnit'] = our_commodityMeasure

# Iterate over the rows of the first DataFrame
# IZMERA PAARVEIDOSANA NO INCH UZ CM, LIELAKA SKAITLA IEVIETOSANA WIDTH, MAZAKA SKAITLA IEVIETOSANA LENGTH, PARVEIDOSANA UZ MUSU KASTU IZMERIEM
# Function to parse box size and update dimensions
def parse_box_size(box_size):
        if box_size in box_sizes:
            return box_sizes[box_size]
        else:
            return None, None, None

# Iterate over the rows of the first DataFrame to set the dimensions
for index, row in df1.iterrows():
    box_size = row['BoxSize']
    length, width, height = parse_box_size(box_size)
    if length and width and height:
        df2.at[index, 'length'] = length
        df2.at[index, 'width'] = width
        df2.at[index, 'height'] = height
    else:
        # If the box size is not found, set default values or handle the error
        df2.at[index, 'length'] = 0
        df2.at[index, 'width'] = 0
        df2.at[index, 'height'] = 0

# Atjaunojam datus jaunaja excel faila
# JAUNA FAILA REDIGESANA PEC DATU PARKOPESANAS
df2.to_excel(destination_file, index=False)

df = pd.read_excel(destination_file)


def split_city_and_state(city_state):
    if isinstance(city_state, str):
        parts = city_state.split(', ')

        if len(parts) == 2:
            city = parts[0].strip()
            state_or_province = parts[1].strip()
            
            if state_or_province in us_states:
                if state_or_province in us_state_to_abbrev:
                    state_or_province = us_state_to_abbrev[state_or_province]
                return city, state_or_province
            elif state_or_province in canadian_provinces:
                if state_or_province in ca_states:
                    state_or_province = canadian_provinces[state_or_province]
                return city, canadian_provinces[state_or_province]
          
    return city_state, None

# Ensure 'carriageValue' and 'customsValue' columns exist and are of type string
if 'carriageValue' not in df.columns:
    df['carriageValue'] = ""

if 'customsValue' not in df.columns:
    df['customsValue'] = ""

df['carriageValue'] = df['carriageValue'].astype(str)
df['customsValue'] = df['customsValue'].astype(str)

# Ensure 'width' and 'length' columns exist in the second DataFrame
# STATA IZNEMSANA NO PILSETAS UN CENAS NOTEIKSANA PEC VALSTS
if 'width' not in df2.columns:
    df2['width'] = 0

if 'length' not in df2.columns:
    df2['length'] = 0

# Ensure 'serviceType' column exists in the second DataFrame and cast to string
if 'serviceType' not in df2.columns:
    df2['serviceType'] = ""
else:
    df2['serviceType'] = df2['serviceType'].astype(str)    

for index, row in df.iterrows():
    city = row['recipientCity']
    state = row['recipientState']
    df.at[index, 'recipientCity'] = city

    if state in us_states:
        if state in us_state_to_abbrev:
            state = us_state_to_abbrev[state]
        df.at[index, 'recipientState'] = state
    elif state in ca_states:
        if state in canadian_provinces:
            state = canadian_provinces[state]
        df.at[index, 'recipientState'] = state
    else:
        df.at[index, 'recipientState'] = ""

    if 'recipientCountry' in df.columns:
        country = row['recipientCountry']
        if country in country_abbreviations:
            df.at[index, 'recipientCountry'] = country_abbreviations[country]
            if df.at[index, 'recipientCountry'] == 'CA':
                df.at[index, 'carriageValue'] = '10'
                df.at[index, 'customsValue'] = '10'
            elif df.at[index, 'recipientCountry'] == 'US':
                df.at[index, 'carriageValue'] = '60'
                df.at[index, 'customsValue'] = '60'    
            else:
                df.at[index, 'carriageValue'] = '40'
                df.at[index, 'customsValue'] = '40'


# Iterate over the rows of the second DataFrame to set the serviceType
# Explicitly cast 'serviceType' column to string to avoid FutureWarning
if 'serviceType' not in df.columns:
    df['serviceType'] = ""
else:
    df['serviceType'] = df['serviceType'].astype(str)

# First, update serviceType to 'FEDEX_INTERNATIONAL_PRIORITY' if it contains the word 'Express'
for index, row in df.iterrows():
    if 'Express' in row['serviceType']:
        df.at[index, 'serviceType'] = 'FEDEX_INTERNATIONAL_PRIORITY'

# SERVISA VEIDA NOTEIKSANA PEC IZMERA UN VALSTS
for index, row in df.iterrows():
    # Skip rows where serviceType is already set to 'FEDEX_INTERNATIONAL_PRIORITY'
    if df.at[index, 'serviceType'] == 'FEDEX_INTERNATIONAL_PRIORITY':
        continue

    width = row['length']
    length = row['width']
    country = row['recipientCountry']
    
    # Use the dictionary to set the serviceType
    if (width, length, country) in service_type_mappings:
        df.at[index, 'serviceType'] = service_type_mappings[(width, length, country)]

# TELEFONA NUMURA IEVEITOSANA JA TUKSUMS
# Ensure 'recipientContactNumber' column exists and fill empty values with 10 zeros
if 'recipientContactNumber' in df.columns:
    df['recipientContactNumber'] = df['recipientContactNumber'].fillna('0000000000')
else:
    df['recipientContactNumber'] = '0000000000'


df.to_excel(destination_file, index=False)

print("Parkopets")