import pandas as pd
import xml.etree.ElementTree as ET

def excel_to_xml(excel_file, xml_file):
    # Load Excel file
    df = pd.read_excel(excel_file)
    
    # Create XML root
    root = ET.Element("openqcat")
    
    # Create header
    header = ET.SubElement(root, "header")
    
    # Add generator_info
    generator_info = ET.SubElement(header, "generator_info")
    generator_info.text = str(df['GENERATOR_INFO'][0])
    
    # Add catalog
    catalog = ET.SubElement(header, "catalog")
    catalog_id = ET.SubElement(catalog, "catalog_id")
    catalog_id.text = str(df['CATALOG_ID'][0])
    catalog_name = ET.SubElement(catalog, "catalog_name")
    catalog_name.text = str(df['CATALOG_NAME'][0])
    language = ET.SubElement(catalog, "language")
    language.text = str(df['LANGUAGE'][0])
    catalog_version = ET.SubElement(catalog, "catalog_version")
    catalog_version.text = str(df['CATALOG_VERSION'][0])
    generation_date = ET.SubElement(catalog, "generation_date")
    generation_date.text = str(df['GENERATION_DATE'][0])
    
    # Add document creator
    document_creator = ET.SubElement(header, "document_creator")
    first_name = ET.SubElement(document_creator, "first_name")
    first_name.text = str(df['FIRST_NAME'][0])
    last_name = ET.SubElement(document_creator, "last_name")
    last_name.text = str(df['LAST_NAME'][0])
    phone = ET.SubElement(document_creator, "phone")
    phone.text = str(df['PHONE'][0])
    
    # Add supplier
    supplier = ET.SubElement(header, "supplier")
    supplier_id = ET.SubElement(supplier, "supplier_id")
    supplier_id.text = str(df['SUPPLIER_ID'][0])
    supplier_name = ET.SubElement(supplier, "supplier_name")
    supplier_name.text = str(df['SUPPLIER_NAME'][0])
    
    address = ET.SubElement(supplier, "address")
    name = ET.SubElement(address, "name")
    name.text = str(df['NAME'][0])
    street = ET.SubElement(address, "street")
    street.text = str(df['STREET'][0])
    zip_code = ET.SubElement(address, "zip")
    zip_code.text = str(df['ZIP'][0])
    city = ET.SubElement(address, "city")
    city.text = str(df['CITY'][0])
    country = ET.SubElement(address, "country")
    country.text = str(df['COUNTRY'][0])
    
    contact = ET.SubElement(supplier, "contact")
    salutation = ET.SubElement(contact, "salutation")
    salutation.text = str(df['SALUTATION'][0])
    contact_first_name = ET.SubElement(contact, "first_name")
    contact_first_name.text = str(df['FIRST_NAME'][0])
    contact_last_name = ET.SubElement(contact, "last_name")
    contact_last_name.text = str(df['LAST_NAME'][0])
    contact_phone = ET.SubElement(contact, "phone")
    contact_phone.text = str(df['PHONE'][0])
    contact_email = ET.SubElement(contact, "email")
    contact_email.text = str(df['EMAIL'][0])
    contact_role = ET.SubElement(contact, "role")
    contact_role.text = str(df['CONTACT_ROLE'][0])
    
    # Iterate over rows to create service elements
    for i, row in df.iterrows():
        if pd.notna(row['SERVICE']):
            service = ET.SubElement(root, "service")
            service_id = ET.SubElement(service, "service_id")
            service_id.text = str(row['SERVICE_ID'])
            service_name = ET.SubElement(service, "service_name")
            service_name.text = row['SERVICE']
            description = ET.SubElement(service, "description")
            description.text = row['DESCRIPTION_LONG']
            start_date = ET.SubElement(service, "start_date")
            start_date.text = str(row['START_DATE'])
            end_date = ET.SubElement(service, "end_date")
            end_date.text = str(row['END_DATE'])
            
            location = ET.SubElement(service, "location")
            location_name = ET.SubElement(location, "name")
            location_name.text = row['LOCATION']
            
            location_address = ET.SubElement(location, "address")
            loc_street = ET.SubElement(location_address, "street")
            loc_street.text = row['STREET']
            loc_city = ET.SubElement(location_address, "city")
            loc_city.text = row['CITY']
            loc_zip = ET.SubElement(location_address, "zip")
            loc_zip.text = row['ZIP']
            loc_country = ET.SubElement(location_address, "country")
            loc_country.text = row['COUNTRY']
            
            price = ET.SubElement(service, "price")
            price.text = str(row['PRICE_AMOUNT'])
            currency = ET.SubElement(service, "currency")
            currency.text = row['PRICE_CURRENCY']
            
            # Add additional fields if they exist in the Excel
            if 'SERVICE_PRICE_DETAILS' in df.columns:
                service_price_details = ET.SubElement(service, "service_price_details")
                service_price_details.text = row['SERVICE_PRICE_DETAILS']
            if 'REMARKS' in df.columns:
                remarks = ET.SubElement(service, "remarks")
                remarks.text = row['REMARKS']
            if 'MIME_INFO' in df.columns and pd.notna(row['MIME_INFO']):
                mime_info = ET.SubElement(service, "mime_info")
                mime_element = ET.SubElement(mime_info, "mime_element")
                mime_source = ET.SubElement(mime_element, "mime_source")
                mime_source.text = row['MIME_SOURCE']
    
    # Write to XML file
    tree = ET.ElementTree(root)
    tree.write(xml_file, encoding='utf-8', xml_declaration=True)

def validate_xml(xml_file):
    try:
        tree = ET.parse(xml_file)
        return True, "XML is valid."
    except ET.ParseError as e:
        return False, f"XML is invalid: {e}"

# Example usage
excel_to_xml('/mnt/data/courses.xlsx', '/mnt/data/output.xml')

# Validate the generated XML file
xml_file_path = '/mnt/data/output.xml'
is_valid, message = validate_xml(xml_file_path)
print(is_valid, message)
