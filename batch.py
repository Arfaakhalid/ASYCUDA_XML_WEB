import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import warnings
import streamlit as st
import zipfile
from io import BytesIO
import tempfile
import uuid
import time
from datetime import datetime
import base64
from pathlib import Path
import glob

# Suppress warnings
warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="Aruba ASYCUDA XML Generator",
    page_icon="🏝️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Consignment-specific fixed values (for LV02 2025 6241)
CONSIGNMENT_VALUES = {
    'total_invoice': '2006.64',
    'total_cif': '4212.99', 
    'total_cost': '621.1',
    'external_freight_foreign': '4.78',
    'external_freight_national': '8.56',
    'insurance_foreign': '0.58',
    'insurance_national': '1.04',
    'other_cost_foreign': '0.47',
    'other_cost_national': '0.84',
    'total_cif_itm': '70.8',
    'statistical_value': '71',
    'alpha_coefficient': '0.0168042100227245',
    'duty_tax_base': '71',
    'duty_tax_rate': '6',
    'duty_tax_amount': '4.3',
    'total_item_taxes': '347.75',
    'calculation_working_mode': '0',
    'container_flag': 'False',
    'delivery_terms_code': 'DDP',
    'currency_rate': '1.79',
    'manifest_reference': 'LV02 2025 6241',
    'total_forms': '16'
}

# Custom CSS with Aruba Theme and Dashboard Style
def set_aruba_theme():
    st.markdown("""
    <style>
    /* Main background with Aruba colors */
    .stApp {
        background: linear-gradient(135deg, #0047AB 0%, #009CDE 50%, #FF671F 100%);
        color: #ffffff;
        min-height: 100vh;
    }
    
    /* Header styling */
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #FFD700;
        text-align: center;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
        font-family: 'Arial Black', sans-serif;
    }
    
    .sub-header {
        font-size: 1.2rem;
        color: #FFFFFF;
        text-align: center;
        margin-bottom: 1.5rem;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
        font-weight: 300;
    }
    
    /* Dashboard container styling */
    .dashboard-container {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 1.5rem;
        margin: 0.5rem 0;
        border: 3px solid #FFD700;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        height: fit-content;
        min-height: 300px;
    }
    
    /* Compact info box */
    .info-box {
        background: linear-gradient(135deg, #0047AB 0%, #009CDE 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #FF671F;
        margin-bottom: 1rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        font-size: 0.9rem;
    }
    
    /* Success box */
    .success-box {
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        color: white;
        padding: 0.8rem;
        border-radius: 8px;
        border-left: 4px solid #FFD700;
        margin-bottom: 0.8rem;
        font-size: 0.9rem;
    }
    
    /* Error box */
    .error-box {
        background: linear-gradient(135deg, #dc3545 0%, #e35d6a 100%);
        color: white;
        padding: 0.8rem;
        border-radius: 8px;
        border-left: 4px solid #FFD700;
        margin-bottom: 0.8rem;
        font-size: 0.9rem;
    }
    
    /* Button styling */
    .stButton button {
        background: linear-gradient(135deg, #FF671F 0%, #FF8C42 100%);
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(255, 103, 31, 0.3);
        width: 100%;
    }
    
    .stButton button:hover {
        background: linear-gradient(135deg, #E55A1B 0%, #FF7B35 100%);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255, 103, 31, 0.4);
    }
    
    /* Secondary button */
    .secondary-button {
        background: linear-gradient(135deg, #6c757d 0%, #5a6268 100%) !important;
    }
    
    /* File uploader styling */
    .upload-section {
        background: rgba(255, 255, 255, 0.1);
        border: 2px dashed #FFD700;
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    
    /* Progress bar styling */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #FF671F 0%, #FFD700 100%);
    }
    
    /* Metric cards */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #0047AB 0%, #009CDE 100%);
        color: white;
        padding: 0.8rem;
        border-radius: 8px;
        border: 2px solid #FFD700;
        text-align: center;
    }
    
    [data-testid="stMetricLabel"] {
        color: #FFD700 !important;
        font-weight: bold;
    }
    
    [data-testid="stMetricValue"] {
        color: white !important;
        font-size: 1.5rem !important;
        font-weight: bold;
    }
    
    [data-testid="stMetricDelta"] {
        color: #FFD700 !important;
        font-weight: bold;
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #0047AB 0%, #009CDE 100%);
        color: white;
        border-radius: 5px;
        font-weight: bold;
    }
    
    /* Log container */
    .log-container {
        background: #1a1a1a;
        color: #00ff00;
        padding: 0.8rem;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        border: 2px solid #FFD700;
        max-height: 200px;
        overflow-y: auto;
        font-size: 0.8rem;
    }
    
    /* File list styling */
    .file-list {
        max-height: 150px;
        overflow-y: auto;
        background: rgba(0, 0, 0, 0.05);
        border-radius: 5px;
        padding: 0.5rem;
        margin: 0.5rem 0;
    }
    
    .file-item {
        padding: 0.3rem;
        margin: 0.2rem 0;
        background: rgba(255, 255, 255, 0.8);
        border-radius: 3px;
        border-left: 3px solid #0047AB;
        font-size: 0.8rem;
    }
    
    /* Remove file size limit warning */
    .stFileUploader > div > small {
        display: none;
    }
    
    /* Make containers responsive */
    @media (max-width: 768px) {
        .main-header {
            font-size: 2rem;
        }
        .sub-header {
            font-size: 1rem;
        }
        .dashboard-container {
            padding: 1rem;
            margin: 0.3rem 0;
        }
    }
    
    /* Custom folder browser styling */
    .folder-browser {
        background: rgba(255, 255, 255, 0.9);
        border: 2px solid #0047AB;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    
    /* Remove default Streamlit container backgrounds */
    .st-emotion-cache-1jicfl2 {
        background: transparent !important;
    }
    .st-emotion-cache-1r6slb0 {
        background: transparent !important;
    }
    </style>
    """, unsafe_allow_html=True)

def read_excel_data(file_content):
    """Read Excel file with exact ASYCUDA structure"""
    sad_data = {}
    items_data = []
    
    try:
        # Read SAD sheet
        try:
            sad_df = pd.read_excel(BytesIO(file_content), sheet_name='SAD')
            if not sad_df.empty:
                sad_row = sad_df.iloc[0]
                for col in sad_df.columns:
                    if pd.notna(sad_row[col]):
                        sad_data[col] = str(sad_row[col])
                    else:
                        sad_data[col] = ''
        except Exception as e:
            st.warning(f"Warning reading SAD sheet: {str(e)}")
        
        # Read Items sheet
        try:
            items_df = pd.read_excel(BytesIO(file_content), sheet_name='Items')
            if not items_df.empty:
                for _, row in items_df.iterrows():
                    item = {}
                    for col in items_df.columns:
                        if pd.notna(row[col]):
                            item[col] = str(row[col])
                        else:
                            item[col] = ''
                    items_data.append(item)
        except Exception as e:
            st.warning(f"Warning reading Items sheet: {str(e)}")
        
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
    
    return sad_data, items_data

def calculate_form_totals(items_data):
    """Calculate form-specific totals (per XML file)"""
    invoice_foreign_total = 0
    for item in items_data:
        inv_foreign = item.get('Invoice Amount_foreign_currency', '0')
        try:
            invoice_foreign_total += float(inv_foreign) if inv_foreign else 0
        except:
            pass
    
    return invoice_foreign_total

def add_element(parent, tag_name, text_content):
    """Helper method to add element with text content"""
    element = ET.SubElement(parent, tag_name)
    if text_content and text_content != 'nan' and text_content != 'None' and text_content != '':
        element.text = str(text_content)
    return element

def create_valuation_subsections(parent, form_invoice_foreign):
    """Create valuation subsections with consignment-specific values"""
    # Invoice
    invoice = ET.SubElement(parent, "Invoice")
    add_element(invoice, "Amount_national_currency", "3591.89")
    add_element(invoice, "Amount_foreign_currency", str(form_invoice_foreign))
    add_element(invoice, "Currency_code", "USD")
    add_element(invoice, "Currency_name", "Geen vreemde valuta")
    add_element(invoice, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # External_freight
    external = ET.SubElement(parent, "External_freight")
    add_element(external, "Amount_national_currency", "509.24")
    add_element(external, "Amount_foreign_currency", "17.27")
    add_element(external, "Currency_code", "USD")
    add_element(external, "Currency_name", "Geen vreemde valuta")
    add_element(external, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Internal_freight
    internal = ET.SubElement(parent, "Internal_freight")
    add_element(internal, "Amount_national_currency", "0")
    add_element(internal, "Amount_foreign_currency", "0")
    add_element(internal, "Currency_code", "")
    add_element(internal, "Currency_name", "Geen vreemde valuta")
    add_element(internal, "Currency_rate", "0")
    
    # Insurance
    insurance = ET.SubElement(parent, "Insurance")
    add_element(insurance, "Amount_national_currency", "62.26")
    add_element(insurance, "Amount_foreign_currency", "1.00875")
    add_element(insurance, "Currency_code", "USD")
    add_element(insurance, "Currency_name", "Geen vreemde valuta")
    add_element(insurance, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Other_cost
    other = ET.SubElement(parent, "Other_cost")
    add_element(other, "Amount_national_currency", "49.6")
    add_element(other, "Amount_foreign_currency", "")
    add_element(other, "Currency_code", "USD")
    add_element(other, "Currency_name", "Geen vreemde valuta")
    add_element(other, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Deduction
    deduction = ET.SubElement(parent, "Deduction")
    add_element(deduction, "Amount_national_currency", "0")
    add_element(deduction, "Amount_foreign_currency", "0")
    add_element(deduction, "Currency_code", "USD")
    add_element(deduction, "Currency_name", "Geen vreemde valuta")
    add_element(deduction, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])

def create_item_supplementary_unit(parent, item_data, unit_num):
    """Create supplementary unit with proper structure for items"""
    supp_unit = ET.SubElement(parent, "Supplementary_unit")
    
    if unit_num == '1':
        add_element(supp_unit, "Supplementary_unit_rank", "")
        add_element(supp_unit, "Supplementary_unit_code", item_data.get('Supplementary_unit_code', 'PCE'))
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_1', 'Aantal Stucks'))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_1', ''))
    elif unit_num == '2':
        add_element(supp_unit, "Supplementary_unit_rank", "2")
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_2', ''))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_2', ''))
    else:  # unit_num == '3'
        add_element(supp_unit, "Supplementary_unit_rank", "3")
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_3', ''))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_3', ''))

def create_item_valuation_subsections(parent, item_data):
    """Create valuation subsections for items with consignment-specific values"""
    # Invoice
    invoice = ET.SubElement(parent, "Invoice")
    add_element(invoice, "Amount_national_currency", "")
    add_element(invoice, "Amount_foreign_currency", item_data.get('Invoice Amount_foreign_currency', ''))
    add_element(invoice, "Currency_code", "USD")
    add_element(invoice, "Currency_name", "Geen vreemde valuta")
    add_element(invoice, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # External_freight (consignment-specific per item)
    external = ET.SubElement(parent, "External_freight")
    add_element(external, "Amount_national_currency", CONSIGNMENT_VALUES['external_freight_national'])
    add_element(external, "Amount_foreign_currency", CONSIGNMENT_VALUES['external_freight_foreign'])
    add_element(external, "Currency_code", "USD")
    add_element(external, "Currency_name", "Geen vreemde valuta")
    add_element(external, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Internal_freight
    internal = ET.SubElement(parent, "Internal_freight")
    add_element(internal, "Amount_national_currency", "0")
    add_element(internal, "Amount_foreign_currency", "")
    add_element(internal, "Currency_code", "")
    add_element(internal, "Currency_name", "Geen vreemde valuta")
    add_element(internal, "Currency_rate", "0")
    
    # Insurance (consignment-specific per item)
    insurance = ET.SubElement(parent, "Insurance")
    add_element(insurance, "Amount_national_currency", CONSIGNMENT_VALUES['insurance_national'])
    add_element(insurance, "Amount_foreign_currency", CONSIGNMENT_VALUES['insurance_foreign'])
    add_element(insurance, "Currency_code", "USD")
    add_element(insurance, "Currency_name", "Geen vreemde valuta")
    add_element(insurance, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Other_cost (consignment-specific per item)
    other = ET.SubElement(parent, "Other_cost")
    add_element(other, "Amount_national_currency", CONSIGNMENT_VALUES['other_cost_national'])
    add_element(other, "Amount_foreign_currency", CONSIGNMENT_VALUES['other_cost_foreign'])
    add_element(other, "Currency_code", "USD")
    add_element(other, "Currency_name", "Geen vreemde valuta")
    add_element(other, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Deduction
    deduction = ET.SubElement(parent, "Deduction")
    add_element(deduction, "Amount_national_currency", "0")
    add_element(deduction, "Amount_foreign_currency", "0")
    add_element(deduction, "Currency_code", "USD")
    add_element(deduction, "Currency_name", "Geen vreemde valuta")
    add_element(deduction, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])

def create_item_element(parent, item_data, item_number):
    """Create individual Item element with consignment-specific values"""
    item = ET.SubElement(parent, "Item")
    
    # Packages section
    packages = ET.SubElement(item, "Packages")
    add_element(packages, "Number_of_packages", item_data.get('Number_of_packages', ''))
    add_element(packages, "Marks1_of_packages", item_data.get('Marks1_of_packages', ''))
    add_element(packages, "Marks2_of_packages", item_data.get('Marks2_of_packages', ''))
    add_element(packages, "Kind_of_packages_code", item_data.get('Kind_of_packages_code', 'STKS'))
    add_element(packages, "Kind_of_packages_name", item_data.get('Kind_of_packages_name', 'Stuks'))
    
    # Tariff section
    tariff = ET.SubElement(item, "Tariff")
    add_element(tariff, "Extended_customs_procedure", item_data.get('Extended_customs_procedure', '4000'))
    add_element(tariff, "National_customs_procedure", item_data.get('National_customs_procedure', '00:00:00'))
    add_element(tariff, "Preference_code", item_data.get('Preference_code', ''))
    
    harmonized = ET.SubElement(tariff, "Harmonized_system")
    add_element(harmonized, "Commodity_code", item_data.get('Commodity_code', ''))
    add_element(harmonized, "Precision_4", item_data.get('Precision_4', ''))
    
    # Three supplementary units
    create_item_supplementary_unit(tariff, item_data, '1')
    create_item_supplementary_unit(tariff, item_data, '2') 
    create_item_supplementary_unit(tariff, item_data, '3')
    
    quota = ET.SubElement(tariff, "Quota")
    add_element(quota, "Quota_code", item_data.get('Quota_code', ''))
    
    # Goods_description
    goods_desc = ET.SubElement(item, "Goods_description")
    add_element(goods_desc, "Country_of_origin_code", item_data.get('Country_of_origin_code', 'US'))
    add_element(goods_desc, "Description_of_goods", item_data.get('Description_of_goods', ''))
    add_element(goods_desc, "Commercial_description", item_data.get('Commercial_description', ''))
    
    # Valuation_item with consignment-specific values
    valuation_item = ET.SubElement(item, "Valuation_item")
    add_element(valuation_item, "Rate_of_adjustment", "1")
    add_element(valuation_item, "Total_cost_itm", "")
    add_element(valuation_item, "Total_cif_itm", CONSIGNMENT_VALUES['total_cif_itm'])
    add_element(valuation_item, "Statistical_value", CONSIGNMENT_VALUES['statistical_value'])
    add_element(valuation_item, "Alpha_coeficient_of_apportionment", CONSIGNMENT_VALUES['alpha_coefficient'])
    
    weight = ET.SubElement(valuation_item, "Weight")
    add_element(weight, "Gross_weight_itm", item_data.get('Gross_weight_itm', '0.5'))
    add_element(weight, "Net_weight_itm", item_data.get('Net_weight_itm', '0.5'))
    
    # Item valuation subsections with consignment-specific values
    create_item_valuation_subsections(valuation_item, item_data)
    
    # Previous_document
    prev_doc = ET.SubElement(item, "Previous_document")
    add_element(prev_doc, "Summary_declaration", item_data.get('Summary_declaration', ''))
    add_element(prev_doc, "Summary_declaration_sl", item_data.get('Summary_declaration_sl', '1'))
    
    # Taxation with consignment-specific values
    taxation = ET.SubElement(item, "Taxation")
    add_element(taxation, "Item_taxes_amount", CONSIGNMENT_VALUES['duty_tax_amount'])
    add_element(taxation, "Item_taxes_mode_of_payment", "1")
    
    tax_line = ET.SubElement(taxation, "Taxation_line")
    add_element(tax_line, "Duty_tax_code", "IR")
    add_element(tax_line, "Duty_tax_base", CONSIGNMENT_VALUES['duty_tax_base'])
    add_element(tax_line, "Duty_tax_rate", CONSIGNMENT_VALUES['duty_tax_rate'])
    add_element(tax_line, "Duty_tax_amount", CONSIGNMENT_VALUES['duty_tax_amount'])
    add_element(tax_line, "Duty_tax_MP", "1")

def create_asycuda_xml(sad_data, items_data, filename):
    """Create exact ASYCUDA XML structure with consignment"""
    # Calculate form-specific values
    form_invoice_foreign = calculate_form_totals(items_data)
    
    # Create root element
    root = ET.Element("ASYCUDA")
    
    # SAD section
    sad = ET.SubElement(root, "SAD")
    
    # Assessment_notice section
    assessment_notice = ET.SubElement(sad, "Assessment_notice")
    add_element(assessment_notice, "Total_item_taxes", CONSIGNMENT_VALUES['total_item_taxes'])
    
    items_taxes = ET.SubElement(assessment_notice, "Items_taxes")
    item_tax = ET.SubElement(items_taxes, "Item_tax")
    add_element(item_tax, "Tax_code", sad_data.get('Tax_code', 'IR'))
    add_element(item_tax, "Tax_description", sad_data.get('Tax_description', 'Invoerrechten'))
    add_element(item_tax, "Tax_amount", CONSIGNMENT_VALUES['total_item_taxes'])
    add_element(item_tax, "Tax_mop", sad_data.get('Tax_mop', '1'))
    
    # Properties section
    properties = ET.SubElement(sad, "Properties")
    add_element(properties, "Sad_flow", sad_data.get('Sad_flow', 'I'))
    
    forms = ET.SubElement(properties, "Forms")
    add_element(forms, "Number_of_the_form", sad_data.get('Number_of_the_form', '1'))
    add_element(forms, "Total_number_of_forms", CONSIGNMENT_VALUES['total_forms'])
    
    add_element(properties, "Selected_page", sad_data.get('Selected_page', '1'))
    
    # Identification section
    identification = ET.SubElement(sad, "Identification")
    add_element(identification, "Manifest_reference_number", CONSIGNMENT_VALUES['manifest_reference'])
    
    office_segment = ET.SubElement(identification, "Office_segment")
    add_element(office_segment, "Customs_clearance_office_code", sad_data.get('Customs_clearance_office_code', 'LV01'))
    add_element(office_segment, "Customs_clearance_office_name", sad_data.get('Customs_clearance_office_name', 'Luchthaven Vracht'))
    
    type_elem = ET.SubElement(identification, "Type")
    add_element(type_elem, "Type_of_declaration", sad_data.get('Type_of_declaration', 'INV'))
    add_element(type_elem, "General_procedure_code", sad_data.get('General_procedure_code', '4'))
    
    # Traders section
    traders = ET.SubElement(sad, "Traders")
    
    exporter = ET.SubElement(traders, "Exporter")
    add_element(exporter, "Exporter_code", sad_data.get('Exporter_code', ''))
    add_element(exporter, "Exporter_name", sad_data.get('Exporter_name', ''))
    
    consignee = ET.SubElement(traders, "Consignee")
    add_element(consignee, "Consignee_code", sad_data.get('Consignee_code', '10026483'))
    add_element(consignee, "Consignee_name", sad_data.get('Consignee_name', 'Dhr. Anthony Martina Paradera 1-H Paradera Paradera Aruba'))
    
    financial_trader = ET.SubElement(traders, "Financial")
    add_element(financial_trader, "Financial_code", sad_data.get('Financial_code', ''))
    add_element(financial_trader, "Financial_name", sad_data.get('Financial_name', ''))
    
    # Declarant section
    declarant = ET.SubElement(sad, "Declarant")
    add_element(declarant, "Declarant_code", sad_data.get('Declarant_code', '1160650'))
    add_element(declarant, "Declarant_name", sad_data.get('Declarant_name', 'Dhr. Victor Hoek Alto Vista 133 Alto Vista Noord/Tanki Leendert Aruba'))
    add_element(declarant, "Declarant_representative", sad_data.get('Declarant_representative', 'Lizandra I. Geerman'))
    
    reference = ET.SubElement(declarant, "Reference")
    add_element(reference, "Year", sad_data.get('Reference Year', '2025'))
    add_element(reference, "Number", sad_data.get('Reference Number', ''))
    
    # General_information section
    general_info = ET.SubElement(sad, "General_information")
    
    country = ET.SubElement(general_info, "Country")
    add_element(country, "Country_first_destination", sad_data.get('Country_first_destination', 'US'))
    add_element(country, "Trading_country", sad_data.get('Trading_country', 'US'))
    add_element(country, "Country_of_origin_name", sad_data.get('Country_of_origin_name', 'Verenigde Staten'))
    
    export = ET.SubElement(country, "Export")
    add_element(export, "Export_country_code", sad_data.get('Export_country_code', 'US'))
    add_element(export, "Export_country_name", sad_data.get('Export_country_name', 'Verenigde Staten'))
    add_element(export, "Export_country_region", sad_data.get('Export_country_region', ''))
    
    destination = ET.SubElement(country, "Destination")
    add_element(destination, "Destination_country_code", sad_data.get('Destination_country_code', 'AW'))
    add_element(destination, "Destination_country_name", sad_data.get('Destination_country_name', 'Aruba'))
    add_element(destination, "Destination_country_region", sad_data.get('Destination_country_region', ''))
    
    add_element(general_info, "Value_details", CONSIGNMENT_VALUES['total_cost'])
    add_element(general_info, "CAP", sad_data.get('CAP', ''))
    
    # Transport section
    transport = ET.SubElement(sad, "Transport")
    add_element(transport, "Container_flag", CONSIGNMENT_VALUES['container_flag'])
    add_element(transport, "Location_of_goods", sad_data.get('Location_of_goods', 'RT-01'))
    add_element(transport, "Location_of_goods_address", sad_data.get('Location_of_goods_address', 'Sabana Berde #75'))
    
    means_transport = ET.SubElement(transport, "Means_of_transport")
    
    departure = ET.SubElement(means_transport, "Departure_arrival_information")
    add_element(departure, "Identity", sad_data.get('Departure_arrival_information Identity', 'COPA AIRLINES'))
    add_element(departure, "Nationality", sad_data.get('Departure_arrival_information Nationality', 'PA'))
    
    border = ET.SubElement(means_transport, "Border_information")
    add_element(border, "Identity", sad_data.get('Border_information Identity', ''))
    add_element(border, "Nationality", sad_data.get('Border_information Nationality', ''))
    add_element(border, "Mode", sad_data.get('Border_information Mode', '4'))
    
    delivery = ET.SubElement(transport, "Delivery_terms")
    add_element(delivery, "Code", CONSIGNMENT_VALUES['delivery_terms_code'])
    add_element(delivery, "Place", sad_data.get('Delivery_terms Place', 'USA'))
    
    border_office = ET.SubElement(transport, "Border_office")
    add_element(border_office, "Code", sad_data.get('Border_office Code', 'LV01'))
    add_element(border_office, "Name", sad_data.get('Border_office Name', 'Luchthaven Vracht'))
    
    place_loading = ET.SubElement(transport, "Place_of_loading")
    add_element(place_loading, "Code", sad_data.get('Place_of_loading Code', 'AWAIR'))
    add_element(place_loading, "Name", sad_data.get('Place_of_loading Name', 'Aeropuerto Reina Beatrix'))
    
    # Financial section
    financial = ET.SubElement(sad, "Financial")
    add_element(financial, "Deffered_payment_reference", sad_data.get('Deffered_payment_reference', ''))
    add_element(financial, "Mode_of_payment", sad_data.get('Mode_of_payment', 'CONTANT'))
    
    fin_trans = ET.SubElement(financial, "Financial_transaction")
    add_element(fin_trans, "Code_1", sad_data.get('Financial_transaction Code_1', '1'))
    add_element(fin_trans, "Code_2", sad_data.get('Financial_transaction Code_1', '1'))
    
    bank = ET.SubElement(financial, "Bank")
    add_element(bank, "Branch", sad_data.get('Bank Branch', ''))
    add_element(bank, "Reference", sad_data.get('Bank Reference', ''))
    
    terms = ET.SubElement(financial, "Terms")
    add_element(terms, "Code", sad_data.get('Terms Code', ''))
    add_element(terms, "Description", sad_data.get('Terms Description', ''))
    
    amounts = ET.SubElement(financial, "Amounts")
    add_element(amounts, "Global_taxes", sad_data.get('Amounts Global_taxes', '0'))
    add_element(amounts, "Totals_taxes", CONSIGNMENT_VALUES['total_item_taxes'])
    
    guarantee = ET.SubElement(financial, "Guarantee")
    add_element(guarantee, "Amount", sad_data.get('Guarantee Amount', '0'))
    
    # Transit section
    transit = ET.SubElement(sad, "Transit")
    add_element(transit, "Result_of_control", sad_data.get('Result_of_control', ''))
    
    # Valuation section
    valuation = ET.SubElement(sad, "Valuation")
    add_element(valuation, "Calculation_working_mode", CONSIGNMENT_VALUES['calculation_working_mode'])
    add_element(valuation, "Total_cost", CONSIGNMENT_VALUES['total_cost'])
    add_element(valuation, "Total_cif", CONSIGNMENT_VALUES['total_cif'])
    
    # Valuation subsections with consignment-specific values
    create_valuation_subsections(valuation, form_invoice_foreign)
    
    total = ET.SubElement(valuation, "Total")
    add_element(total, "Total_invoice", CONSIGNMENT_VALUES['total_invoice'])
    add_element(total, "Total_weight", str(len(items_data)))
    
    # Items section
    items_elem = ET.SubElement(root, "Items")
    
    # Create items from Items data with consignment-specific values
    for i, item_data in enumerate(items_data):
        create_item_element(items_elem, item_data, i+1)
    
    return root

def prettify_xml(elem):
    """Convert XML to pretty formatted string"""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def convert_excel_to_xml(file_content, filename):
    """Convert single Excel file to ASYCUDA XML"""
    try:
        # Read data from Excel
        sad_data, items_data = read_excel_data(file_content)
        
        if not sad_data and not items_data:
            return False, f"No valid data found in {filename}"
        
        # Create XML structure
        xml_root = create_asycuda_xml(sad_data, items_data, filename)
        
        # Generate XML content
        xml_content = prettify_xml(xml_root)
        
        return True, xml_content
        
    except Exception as e:
        return False, f"{filename} | Error: {str(e)}"

def main():
    # Set Aruba theme
    set_aruba_theme()
    
    # Header
    st.markdown('<div class="main-header">🏝️ Aruba ASYCUDA XML Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Professional Excel to ASYCUDA XML Conversion • Consignment LV02 2025 6241</div>', unsafe_allow_html=True)
    
    # Initialize session state for files
    if 'all_files' not in st.session_state:
        st.session_state.all_files = []
    
    # Dashboard Layout - Side by Side
    col1, col2 = st.columns([1, 1], gap="medium")
    
    with col1:
        # File Selection Section
        st.header("📁 File Selection")
        
        # File Selection Methods in Tabs
        tab1, tab2 = st.tabs(["📄 Individual Files", "📁 Upload Folder"])
        
        with tab1:
            st.subheader("Select Individual Excel Files")
            individual_files = st.file_uploader(
                "Choose Excel files",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="individual_files",
                help="Select multiple Excel files for conversion"
            )
            
            if individual_files:
                st.success(f"✅ {len(individual_files)} individual file(s) selected")
        
        with tab2:
            st.subheader("Upload Folder with Excel Files")
            st.info("💡 Select multiple Excel files from your folder (works on both local and cloud)")
            
            # Multiple file selection for folder upload
            folder_files = st.file_uploader(
                "Select ALL Excel files from your folder",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="folder_files",
                help="Hold Ctrl/Cmd to select multiple files, or drag and drop all files from your folder"
            )
            
            if folder_files:
                st.success(f"✅ {len(folder_files)} file(s) selected from folder")
                st.info(f"📁 Folder upload complete! Found {len(folder_files)} Excel files")
        
        # Combine all files
        all_files = []
        if individual_files:
            all_files.extend(individual_files)
        if folder_files:
            all_files.extend(folder_files)
        
        # Remove duplicates
        unique_files = []
        seen_files = set()
        for file in all_files:
            file_id = (file.name, getattr(file, 'size', 0))
            if file_id not in seen_files:
                seen_files.add(file_id)
                unique_files.append(file)
        
        # Update session state
        st.session_state.all_files = unique_files
        
        # Display file summary
        if st.session_state.all_files:
            st.markdown("---")
            st.subheader("📋 Selected Files Summary")
            
            # File management buttons
            
            
            with st.expander("View File Details", expanded=True):
                total_size = 0
                for i, file in enumerate(st.session_state.all_files[:15]):  # Show first 15
                    file_size = getattr(file, 'size', 0)
                    total_size += file_size
                    size_mb = file_size / (1024 * 1024) if file_size > 0 else 0
                    st.write(f"**{i+1}.** `{file.name}` ({size_mb:.1f} MB)")
                
                if len(st.session_state.all_files) > 15:
                    st.write(f"... and {len(st.session_state.all_files) - 15} more files")
                
                st.write(f"**Total Size:** {total_size / (1024 * 1024):.1f} MB")
        
        else:
            st.info("📝 No files selected yet. Use the tabs above to select files.")
    
    with col2:
        # Conversion Control Section
        st.header("⚙️ Conversion Control")
        
        # Info box
        st.markdown("""
        <div class="info-box">
        <strong> ASYCUDA XML Generator v4.0</strong><br>
        • Exact Aruba ASYCUDA XML compliance<br>
        • Batch conversion support<br>
        • Real-time progress tracking<br>
        • Made by Arfa Rumman Khalid<br>
        <strong>Consignment:</strong> LV02 2025 6241
        </div>
        """, unsafe_allow_html=True)
        
        # Conversion button
        if st.session_state.all_files:
            if st.button("✅ START CONVERSION", use_container_width=True, type="primary"):
                st.session_state.conversion_started = True
            else:
                st.session_state.conversion_started = False
        else:
            st.button("✅ START CONVERSION", use_container_width=True, disabled=True)
            st.warning("Please select files to enable conversion")
        
        # Conversion results area
        if st.session_state.get('conversion_started', False) and st.session_state.all_files:
            st.markdown("---")
            st.subheader("🔄 Conversion Progress")
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            successful_conversions = 0
            failed_conversions = 0
            conversion_log = []
            
            # Create zip file in memory
            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                total_files = len(st.session_state.all_files)
                
                for i, file in enumerate(st.session_state.all_files):
                    # Update progress
                    progress = (i / total_files)
                    progress_bar.progress(progress)
                    status_text.text(f"🔄 Processing {i+1}/{total_files}: {file.name}")
                    
                    # Convert file
                    try:
                        file_content = file.read()
                        success, result = convert_excel_to_xml(file_content, file.name)
                        
                        if success:
                            xml_filename = file.name.rsplit('.', 1)[0] + '.xml'
                            zip_file.writestr(xml_filename, result)
                            successful_conversions += 1
                            conversion_log.append(f"✅ SUCCESS: {file.name}")
                        else:
                            error_filename = file.name + '_ERROR.txt'
                            zip_file.writestr(error_filename, f"Conversion failed: {result}")
                            failed_conversions += 1
                            conversion_log.append(f"❌ FAILED: {file.name} - {result}")
                            
                    except Exception as e:
                        error_filename = file.name + '_ERROR.txt'
                        zip_file.writestr(error_filename, f"Unexpected error: {str(e)}")
                        failed_conversions += 1
                        conversion_log.append(f"💥 ERROR: {file.name} - {str(e)}")
                    
                    # Small delay for smooth progress animation
                    time.sleep(0.1)
                
                # Final progress update
                progress_bar.progress(1.0)
                status_text.text("✅ Conversion completed!")
            
            zip_buffer.seek(0)
            
            # Results summary
            st.markdown("---")
            st.subheader("📊 Conversion Results")
            
            # Metrics in columns
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Files", total_files)
            with col2:
                st.metric("Successful", successful_conversions)
            with col3:
                st.metric("Failed", failed_conversions)
            
            # Success rate
            success_rate = (successful_conversions / total_files * 100) if total_files > 0 else 0
            st.metric("Success Rate", f"{success_rate:.1f}%")
            
            # Conversion log
            with st.expander("View Conversion Log", expanded=True):
                log_content = "\n".join(conversion_log)
                st.text_area("Conversion Log", log_content, height=150, key="conversion_log")
            
            # Download button
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 DOWNLOAD ASYCUDA XML FILES (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"ASYCUDA_XML_Output_{timestamp}.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary"
            )
            
            # Clear files after successful download
            if st.button("🔄 Start New Conversion", use_container_width=True):
                st.session_state.all_files = []
                st.session_state.conversion_started = False
                st.rerun()

if __name__ == "__main__":
    main()