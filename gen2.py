import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
import re
import pytz

# Load the provided Excel file
file_path = 'DELA-EXCEL6k.xlsx'
df = pd.read_excel(file_path)
df = df.applymap(lambda x: str(x) if not isinstance(x, str) else x)


def format_date_european_to_iso_with_timezone(date_str, timezone_offset='+01:00'):
    try:
        # Parse the European date format (DD.MM.YYYY)
        date = datetime.strptime(date_str, '%d.%m.%Y')
        
        # Format to ISO 8601 with a fixed timezone offset
        formatted_date = date.strftime('%Y-%m-%dT%H:%M:%S.000') + timezone_offset
        return formatted_date
    except ValueError:
        return ""  # Return an empty string if parsing fails




""""
def format_phone_number(number):
    if pd.isna(number):
        return ""
    # Remove any characters that are not digits or plus
    number = re.sub(r"[^\d\+]", "", number)
    # Add '+' if it doesn't start with it, assuming the numbers are all in international format without the '+'
    if not number.startswith('+'):
        number = '+' + number
    return number

def safe_attr_type(value, default_type):
    if pd.isna(value) or value == 'nan':
        return default_type
    try:
        # Assuming the type must be a non-floating integer as per the schema
        int_value = int(float(value))
        return str(int_value)
    except ValueError:
        return default_type



df['PHONE'] = df['PHONE'].apply(format_phone_number)
df['MOBILE'] = df['MOBILE'].apply(format_phone_number)
df['FAX'] = df['FAX'].apply(format_phone_number)
"""


df['START_DATE'] = df['START_DATE'].apply(format_date_european_to_iso_with_timezone)
start_value = df['START_DATE']

df['END_DATE'] = df['END_DATE'].apply(format_date_european_to_iso_with_timezone)
end_value = df['END_DATE']



def create_service_element_updated(row):
    service = ET.Element('SERVICE', mode="new")

    product_id = ET.SubElement(service, 'PRODUCT_ID')
    product_id.text = str(row['PRODUCT_ID'])

    course_type = ET.SubElement(service, 'COURSE_TYPE')
    course_type.text = str(row['COURSE_TYPE'])

    supplier_id_ref = ET.SubElement(service, 'SUPPLIER_ID_REF', type="supplier_specific")
    supplier_id_ref.text = str(row['SUPPLIER_ID_REF'])

    service_details = ET.SubElement(service, 'SERVICE_DETAILS')
    
    title = ET.SubElement(service_details, 'TITLE')
    title.text = row['TITLE'] if not pd.isna(row['TITLE']) else ""

    description_long = ET.SubElement(service_details, 'DESCRIPTION_LONG')
    description_long.text = row['DESCRIPTION_LONG'] if not pd.isna(row['DESCRIPTION_LONG']) else ""

    supplier_alt_pid = ET.SubElement(service_details, 'SUPPLIER_ALT_PID')
    supplier_alt_pid.text = str(row['SUPPLIER_ALT_PID'])

    contact = ET.SubElement(service_details, 'CONTACT')
    contact_role = ET.SubElement(contact, 'CONTACT_ROLE', type="1")
    contact_role.text = row['CONTACT_ROLE'] if not pd.isna(row['CONTACT_ROLE']) else ""
    
    salutation = ET.SubElement(contact, 'SALUTATION')
    salutation.text = row['SALUTATION'] if not pd.isna(row['SALUTATION']) else ""
    
    first_name = ET.SubElement(contact, 'FIRST_NAME')
    first_name.text = row['FIRST_NAME'] if not pd.isna(row['FIRST_NAME']) else ""
    
    last_name = ET.SubElement(contact, 'LAST_NAME')
    last_name.text = row['LAST_NAME'] if not pd.isna(row['LAST_NAME']) else ""
    
    phone = ET.SubElement(contact, 'PHONE')
    phone.text = "+49.211.91382910"
    
    mobile = ET.SubElement(contact, 'MOBILE')
    mobile.text = "+49.211.91382910"
    
    fax = ET.SubElement(contact, 'FAX')
    fax.text = "+49.211.91382939"
    
    emails = ET.SubElement(contact, 'EMAILS')
    email = ET.SubElement(emails, 'EMAIL')
    email.text = row['EMAIL'] if not pd.isna(row['EMAIL']) else ""
    
    url = ET.SubElement(contact, 'URL')
    url.text = row['URL'] if not pd.isna(row['URL']) else ""
    
    id_db = ET.SubElement(contact, 'ID_DB')
    id_db.text = str(row['ID_DB'])

    contact_remarks = ET.SubElement(contact, 'CONTACT_REMARKS')
    contact_remarks.text = row['CONTACT_REMARKS'] if not pd.isna(row['CONTACT_REMARKS']) else ""

    requirements = ET.SubElement(service_details, 'REQUIREMENTS')
    requirements.text = row['REQUIREMENTS'] if not pd.isna(row['REQUIREMENTS']) else ""

    service_date = ET.SubElement(service_details, 'SERVICE_DATE')
    start_date = ET.SubElement(service_date, 'START_DATE')
    start_date.text = "2024-07-01T00:00:00.000+01:00"  
    end_date = ET.SubElement(service_date, 'END_DATE')
    end_date.text =  "2025-03-02T00:00:00.000+01:00"

    keywords = row['KEYWORD'].split(',') if not pd.isna(row['KEYWORD']) else []
    for keyword in keywords:
        keyword_element = ET.SubElement(service_details, 'KEYWORD')
        keyword_element.text = keyword

    target_group = ET.SubElement(service_details, 'TARGET_GROUP')
    target_group_text = ET.SubElement(target_group, 'TARGET_GROUP_TEXT')
    target_group_text.text = row['TARGET_GROUP_TEXT'] if not pd.isna(row['TARGET_GROUP_TEXT']) else ""

    terms_conditions = ET.SubElement(service_details, 'TERMS_AND_CONDITIONS')
    terms_conditions.text = ""

    service_module = ET.SubElement(service_details, 'SERVICE_MODULE')
    education = ET.SubElement(service_module, 'EDUCATION', type="true")

    course_id = ET.SubElement(education, 'COURSE_ID')
    course_id.text = str(row['COURSE_ID'])

    degree = ET.SubElement(education, 'DEGREE', type=str(row['Degree_TYPE1']))
    degree_title = ET.SubElement(degree, 'DEGREE_TITLE')
    degree_title.text = row['DEGREE_TITLE'] if not pd.isna(row['DEGREE_TITLE']) else ""
    degree_exam = ET.SubElement(degree, 'DEGREE_EXAM', type=row['Degree_TYPE2'] if not pd.isna(row['Degree_TYPE2']) else "")
    examiner = ET.SubElement(degree_exam, 'EXAMINER')
    examiner.text = row['EXAMINER'] if not pd.isna(row['EXAMINER']) else "Kein Angabe"
    degree_add_qualification = ET.SubElement(degree, 'DEGREE_ADD_QUALIFICATION')
    degree_add_qualification.text = row['DEGREE_ADD_QUALIFICATION'] if not pd.isna(row['DEGREE_ADD_QUALIFICATION']) else ""
    degree_entitled = ET.SubElement(degree, 'DEGREE_ENTITLED')
    degree_entitled.text = row['DEGREE_ENTITLED'] if not pd.isna(row['DEGREE_ENTITLED']) else ""

    subsidy = ET.SubElement(education, 'SUBSIDY')
    subsidy_description = ET.SubElement(subsidy, 'SUBSIDY_DESCRIPTION')
    subsidy_description.text = row['SUBSIDY_DESCRIPTION'] if not pd.isna(row['SUBSIDY_DESCRIPTION']) else ""

    certificate = ET.SubElement(education, 'CERTIFICATE')
    certificate_status = ET.SubElement(certificate, 'CERTIFICATE_STATUS')
    certificate_status.text = str(row['CERTIFICATE_STATUS'])
    certifier_number = ET.SubElement(certificate, 'CERTIFIER_NUMBER')
    certifier_number.text = str(row['CERTIFIER_NUMBER'])
    cert_validity = ET.SubElement(certificate, 'CERT_VALIDITY')
    

    extended_info = ET.SubElement(education, 'EXTENDED_INFO')
    institution = ET.SubElement(extended_info, 'INSTITUTION', type=str(row['INSTITUTION_TYPE']))
    institution.text = row['INSTITUTION'] if not pd.isna(row['INSTITUTION']) else ""
    instruction_form = ET.SubElement(extended_info, 'INSTRUCTION_FORM', type=str(row['INSTRUCTION_TYPE1']))
    instruction_form.text = row['INSTRUCTION_FORM'] if not pd.isna(row['INSTRUCTION_FORM']) else ""
    instruction_form_name = ET.SubElement(extended_info, 'INSTRUCTION_FORM_NAME')
    instruction_form_name.text = row['INSTRUCTION_FORM_NAME'] if not pd.isna(row['INSTRUCTION_FORM_NAME']) else ""
    instruction_time = ET.SubElement(extended_info, 'INSTRUCTION_TIME', type=str(row['INSTRUCTION_TYPE2']))
    instruction_time.text = row['INSTRUCTION_TIME'] if not pd.isna(row['INSTRUCTION_TIME']) else ""
    inhouse_seminar = ET.SubElement(extended_info, 'INHOUSE_SEMINAR')
    inhouse_seminar.text = str(row['INHOUSE_SEMINAR']).lower() if not pd.isna(row['INHOUSE_SEMINAR']) else ""
    extra_occupational = ET.SubElement(extended_info, 'EXTRA_OCCUPATIONAL')
    extra_occupational.text = str(row['EXTRA_OCCUPATIONAL']).lower() if not pd.isna(row['EXTRA_OCCUPATIONAL']) else ""
    practical_part = ET.SubElement(extended_info, 'PRACTICAL_PART')
    practical_part.text = str(row['PRACTICAL_PART']).lower() if not pd.isna(row['PRACTICAL_PART']) else ""
    education_type = ET.SubElement(extended_info, 'EDUCATION_TYPE', type=str(row['EDUCATION_TYPE2']))
    education_type.text = row['EDUCATION_TYPE1'] if not pd.isna(row['EDUCATION_TYPE1']) else ""

    module_course = ET.SubElement(education, 'MODULE_COURSE')
    location = ET.SubElement(module_course, 'LOCATION')
    location_name = ET.SubElement(location, 'NAME')
    location_name.text = row['LOCATION_NAME'] if not pd.isna(row['LOCATION_NAME']) else ""
    location_name2 = ET.SubElement(location, 'NAME2')
    location_name2.text = row['LOCATION_NAME2'] if not pd.isna(row['LOCATION_NAME2']) else ""
    location_street = ET.SubElement(location, 'STREET')
    location_street.text = row['STREET'] if not pd.isna(row['STREET']) else ""
    location_zip = ET.SubElement(location, 'ZIP')
    location_zip.text = row['ZIP'] if not pd.isna(row['ZIP']) else ""
    location_zipbox = ET.SubElement(location, 'ZIPBOX')
    location_zipbox.text = row['ZIPBOX'] if not pd.isna(row['ZIPBOX']) else ""
    location_city = ET.SubElement(location, 'CITY')
    location_city.text = row['CITY'] if not pd.isna(row['CITY']) else ""
    
    location_district = ET.SubElement(location, 'DISTRICT')
    location_district.text = row['DISTRICT'] if not pd.isna(row['DISTRICT']) else ""
    
    location_state = ET.SubElement(location, 'STATE')
    location_state.text = row['STATE'] if not pd.isna(row['STATE']) else ""
    
    location_country = ET.SubElement(location, 'COUNTRY')
    location_country.text = row['COUNTRY'] if not pd.isna(row['COUNTRY']) else ""
    
    location_phone = ET.SubElement(location, 'PHONE')
    location_phone.text = "+49.211.91382910"
    
    location_mobile = ET.SubElement(location, 'MOBILE')
    location_mobile.text = "+49.211.91382910"
    
    location_fax = ET.SubElement(location, 'FAX')
    location_fax.text = "+49.211.91382939"
    
    location_emails = ET.SubElement(location, 'EMAILS')
    location_email = ET.SubElement(location_emails, 'EMAIL')
    location_email.text = row['EMAIL'] if not pd.isna(row['EMAIL']) else ""
    
    location_url = ET.SubElement(location, 'URL')
    location_url.text = row['URL'] if not pd.isna(row['URL']) else ""
    
    location_address_remarks = ET.SubElement(location, 'ADDRESS_REMARKS')
    location_address_remarks.text = row['ADDRESS_REMARKS'] if not pd.isna(row['ADDRESS_REMARKS']) else ""
    
    duration = ET.SubElement(module_course, 'DURATION', type=str(row['DURATION_TYPE']))
    
    instruction_remarks = ET.SubElement(module_course, 'INSTRUCTION_REMARKS')
    instruction_remarks.text = row['INSTRUCTION_REMARKS'] if not pd.isna(row['INSTRUCTION_REMARKS']) else ""
    
    flexible_start = ET.SubElement(module_course, 'FLEXIBLE_START')
    flexible_start.text = str(row['FLEXIBLE_START']).lower() if not pd.isna(row['FLEXIBLE_START']) else ""
    
    extended_info = ET.SubElement(module_course, 'EXTENDED_INFO')
    segment_type = ET.SubElement(extended_info, 'SEGMENT_TYPE', type=str(row['SEGMENT_TYPE2']))
    
    announcement = ET.SubElement(service_details, 'ANNOUNCEMENT')
    
    announcement_start_date = ET.SubElement(announcement, 'START_DATE')
    announcement_start_date.text = "2024-07-01+01:00"
    announcement_end_date = ET.SubElement(announcement, 'END_DATE')
    announcement_end_date.text = "2025-03-02+01:00"
    
    
    service_classification = ET.SubElement(service, 'SERVICE_CLASSIFICATION')
    reference_classification_system_name = ET.SubElement(service_classification, 'REFERENCE_CLASSIFICATION_SYSTEM_NAME')
    reference_classification_system_name.text = row['REFERENCE_CLASSIFICATION_SYSTEM_NAME'] if not pd.isna(row['REFERENCE_CLASSIFICATION_SYSTEM_NAME']) else ""
    
    feature = ET.SubElement(service_classification, 'FEATURE')
    fname = ET.SubElement(feature, 'FNAME')
    fname.text = row['FNAME'] if not pd.isna(row['FNAME']) else ""
    fvalue = ET.SubElement(feature, 'FVALUE')
    fvalue.text = row['FVALUE'] if not pd.isna(row['FVALUE']) else ""
    
    service_price_details = ET.SubElement(service, 'SERVICE_PRICE_DETAILS')
    service_price = ET.SubElement(service_price_details, 'SERVICE_PRICE')
    price_amount = ET.SubElement(service_price, 'PRICE_AMOUNT')
    price_amount.text = str(row['PRICE_AMOUNT']) if not pd.isna(row['PRICE_AMOUNT']) else ""
    price_currency = ET.SubElement(service_price, 'PRICE_CURRENCY')
    price_currency.text = row['PRICE_CURRENCY'] if not pd.isna(row['PRICE_CURRENCY']) else ""
    
    mime_info = ET.SubElement(service, 'MIME_INFO')
    mime_element = ET.SubElement(mime_info, 'MIME_ELEMENT')
    mime_source = ET.SubElement(mime_element, 'MIME_SOURCE')
    mime_source.text = row['MIME_SOURCE'] if not pd.isna(row['MIME_SOURCE']) else ""
    
    return service

def dataframe_to_xml(df):

    root = ET.Element('NEW_CATALOG', FULLCATALOG="true")

    

    for index, row in df.iterrows():

        service = create_service_element_updated(row)

        root.append(service)

    

    tree = ET.ElementTree(root)

    return tree



# Generate the XML tree from the new DataFrame

xml_tree_new = dataframe_to_xml(df)



# Save the new XML to a file

output_file_new = 'Delafinal6k.xml'
with open(output_file_new, 'wb') as f:
    xml_tree_new.write(f, encoding='utf-8', xml_declaration=True)

