import streamlit as st
from pypdf import PdfReader
import pandas as pd
from io import BytesIO
import re
import csv
from datetime import datetime

def load_product_catalog():
    """Load product catalog from Excel file"""
    try:
        # Load the Excel file
        catalog_df = pd.read_excel('ProductCatalog.xlsx')
        
        # Create a dictionary mapping product names to their codes and sizes
        product_catalog = {}
        for _, row in catalog_df.iterrows():
            # Get values using the exact column names
            product_name = str(row['Product Name']).strip()
            product_code = str(row['Code']).strip()
            size = str(row['Size']).strip()
            
            # Skip empty product names
            if not product_name:
                continue
            
            product_catalog[product_name] = {
                'Product Code': product_code,
                'Size': size
            }
        
        return product_catalog
    except Exception as e:
        st.error(f"Error loading product catalog: {str(e)}")
        return {}

def extract_pdf_data(pdf_path):
    try:
        # Load product catalog
        product_catalog = load_product_catalog()
        
        # Read the entire file into memory first
        with open(pdf_path, 'rb') as file:
            pdf_data = file.read()
        
        # Create a BytesIO object from the file data
        pdf_stream = BytesIO(pdf_data)
        
        # Create PDF reader
        pdf_reader = PdfReader(pdf_stream)
        
        # Extract form fields
        form_fields = pdf_reader.get_fields()
        
        # Debug: Display all field names and values
        st.write("Available PDF Form Fields:")
        for field_name, field in form_fields.items():
            value = field.get('/V', '')
            st.write(f"Field: {field_name}, Value: {value}")
        
        # Process the form data
        order_data = {}
        line_items = []
        
        # Extract customer information
        customer_info = {}
        # Add more potential customer field variations
        customer_fields = [
            'Customer Name', 'Company', 'Email', 'Phone',
            'customer_name', 'company', 'email', 'phone',
            'name', 'customer', 'contact'
        ]
        for field in customer_fields:
            if field in form_fields:
                customer_info[field] = form_fields[field].get('/V', '')
        
        # Extract line items
        # Look for fields containing "Order" or "order" in any position
        order_patterns = ['Order', 'order']
        
        # Add specific product field name patterns
        specific_products = {
            # Classic Pumpkin variants
            'Classic Pumpkin': {
                '550': 'CCPUM - Classic Pumpkin - 550g x 4',
                '350': 'CPCPUM - Classic Pumpkin - 350g x 6'
            },
            # Sweet Potato variant
            'Sweet Potato': {
                '550': 'CSP - Sweet Potato - 550g x 4'
            }
        }

        # Process each field
        for field_name, field in form_fields.items():
            value = field.get('/V', '')
            try:
                quantity = int(value)
                if quantity > 0:
                    st.write(f"Processing field: {field_name}")
                    
                    # Direct mapping of field names to product details
                    field_mappings = {
                        # Soup products
                        'Ancient Grain': {
                            'product_name': 'CAGR - Ancient Grain - 550g x 4',
                            'code': 'CAGR',
                            'size': '550g x 4'
                        },
                        'Beef & Barley': {
                            'product_name': 'CBAB - Beef & Barley - 550g x 5',
                            'code': 'CBAB',
                            'size': '550g x 5'
                        },
                        'Chicken & Corn': {
                            'product_name': 'CCCORN - Chicken & Corn - 550g x 4',
                            'code': 'CCCORN',
                            'size': '550g x 4'
                        },
                        'Chicken Noodle': {
                            'product_name': 'CNCH - Chicken Noodle - 550g x 4',
                            'code': 'CNCH',
                            'size': '550g x 4'
                        },
                        'Chunky Minestrone': {
                            'product_name': 'CCMIN - Chunky Minestrone - 550g x 4',
                            'code': 'CCMIN',
                            'size': '550g x 4'
                        },
                        'Creamy Chicken': {
                            'product_name': 'CCRC - Creamy Chicken - 550g x 4',
                            'code': 'CCRC',
                            'size': '550g x 4'
                        },
                        'Creamy Mushroom': {
                            'product_name': 'CMUSH - Creamy Mushroom - 550g x 4',
                            'code': 'CMUSH',
                            'size': '550g x 4'
                        },
                        'Creamy Pumpkin': {
                            'product_name': 'CRMPUM - Creamy Pumpkin - 550g x 4',
                            'code': 'CRMPUM',
                            'size': '550g x 4'
                        },
                        'French Onion': {
                            'product_name': 'CFON - French Onion - 550g x 4',
                            'code': 'CFON',
                            'size': '550g x 4'
                        },
                        'Hearty Beef': {
                            'product_name': 'CHB - Hearty Beef - 550g x 4',
                            'code': 'CHB',
                            'size': '550g x 4'
                        },
                        'Hearty Chicken': {
                            'product_name': 'CHCH - Hearty Chicken - 550g x 4',
                            'code': 'CHCH',
                            'size': '550g x 4'
                        },
                        'Hearty Vegetable': {
                            'product_name': 'CHVS - Hearty Vegetable - 550g x 4',
                            'code': 'CHVS',
                            'size': '550g x 4'
                        },
                        'Lentil': {
                            'product_name': 'CLEN - Lentil - 550g x 4',
                            'code': 'CLEN',
                            'size': '550g x 4'
                        },
                        'Pea & Ham': {
                            'product_name': 'CPHAM - Pea & Ham - 550g x 4',
                            'code': 'CPHAM',
                            'size': '550g x 4'
                        },
                        'Potato & Leek': {
                            'product_name': 'CPOTL - Potato & Leek - 550g x 4',
                            'code': 'CPOTL',
                            'size': '550g x 4'
                        },
                        'Tomato': {
                            'product_name': 'CTOM - Tomato - 350g x 6',
                            'code': 'CTOM',
                            'size': '350g x 6'
                        },
                        
                        # 350g variants
                        'Chixken & Vegetable 350': {
                            'product_name': 'CCVEG - Chicken & Vegetable - 350g x 6',
                            'code': 'CCVEG',
                            'size': '350g x 6'
                        },
                        'Tomato 350': {
                            'product_name': 'CTOM - Tomato - 350g x 6',
                            'code': 'CTOM',
                            'size': '350g x 6'
                        },
                        
                        # 600g variants
                        'Asian Chicken 600': {
                            'product_name': 'FWACHC - Asian Chicken - 600g x 4',
                            'code': 'FWACHC',
                            'size': '600g x 4'
                        },
                        'Crab & Corn 600': {
                            'product_name': 'CRCC - Crab & Corn - 600g x 4',
                            'code': 'CRCC',
                            'size': '600g x 4'
                        },
                        'Thai Prawn 600': {
                            'product_name': 'FWTPRA - Thai Prawn - 600g x 4',
                            'code': 'FWTPRA',
                            'size': '600g x 4'
                        },
                        'Moroccan Harira 600': {
                            'product_name': 'FWMH - Moroccan Harira - 600g x 4',
                            'code': 'FWMH',
                            'size': '600g x 4'
                        },
                        
                        # Curry products
                        'Chicken Tagine': {
                            'product_name': 'CKMCT - Chicken Tagine - 500g x 4',
                            'code': 'CKMCT',
                            'size': '500g x 4'
                        },
                        'Red Chicken Curry': {
                            'product_name': 'CKMRCC - Red Chicken Curry - 500g x 4',
                            'code': 'CKMRCC',
                            'size': '500g x 4'
                        },
                        'Red Vegetable Curry': {
                            'product_name': 'CKMRVC - Red Vegetable Curry - 500g x 4',
                            'code': 'CKMRVC',
                            'size': '500g x 4'
                        },
                        
                        # Hummus products
                        'Baba Ganoush': {
                            'product_name': 'CBG - Baba Ganoush - 200g x 6',
                            'code': 'CBG',
                            'size': '200g x 6'
                        },
                        'Beetroot Hummus': {
                            'product_name': 'CBHUM - Beetroot Hummus - 200g x 6',
                            'code': 'CBHUM',
                            'size': '200g x 6'
                        },
                        'Chilli Kalamata Hummus': {
                            'product_name': 'CCKHU - Chilli Kalamata Hummus - 200g x 6',
                            'code': 'CCKHU',
                            'size': '200g x 6'
                        },
                        'Harissa Hummus': {
                            'product_name': 'CHHU - Harissa Hummus - 200g x 6',
                            'code': 'CHHU',
                            'size': '200g x 6'
                        },
                        'Mediterranean Hummus': {
                            'product_name': 'CMHU - Mediterranean Hummus - 200g x 6',
                            'code': 'CMHU',
                            'size': '200g x 6'
                        },
                        'Olive Hummus': {
                            'product_name': 'COHU - Olive Hummus - 200g x 6',
                            'code': 'COHU',
                            'size': '200g x 6'
                        },
                        'Pine Nut Hummus': {
                            'product_name': 'CPNHU - Pine Nut Hummus - 200g x 6',
                            'code': 'CPNHU',
                            'size': '200g x 6'
                        },
                        'Roasted Garlic Hummus': {
                            'product_name': 'CRGHU - Roasted Garlic Hummus - 200g x 6',
                            'code': 'CRGHU',
                            'size': '200g x 6'
                        },
                        
                        # Other dips
                        'Hummus 1kg': {
                            'product_name': 'CHUM1 - Hummus - 1kg x 1',
                            'code': 'CHUM1',
                            'size': '1kg x 1'
                        },
                        'Garlic Dip 180': {
                            'product_name': 'CDIP180 - Garlic Dip - 180g x 1',
                            'code': 'CDIP180',
                            'size': '180g x 1'
                        },
                        'Taramosalata 180': {
                            'product_name': 'CTAR180 - Taramosalata - 180g x 1',
                            'code': 'CTAR180',
                            'size': '180g x 1'
                        },
                        'Smoked Salmon 170': {
                            'product_name': 'CSSM170 - Smoked Salmon - 170g x 1',
                            'code': 'CSSM170',
                            'size': '170g x 1'
                        }
                    }
                    
                    # Check if this is a mapped field
                    if field_name in field_mappings:
                        mapping = field_mappings[field_name]
                        product_name = mapping['product_name']
                        code = mapping['code']
                        size = mapping['size']
                        st.write(f"Matched specific product: {product_name}")
                    else:
                        # Handle Classic Pumpkin variants
                        if 'Classic Pumpkin' in field_name:
                            if '350' in field_name:
                                product_name = 'CPCPUM - Classic Pumpkin - 350g x 6'
                                code = 'CPCPUM'
                                size = '350g x 6'
                            else:
                                product_name = 'CCPUM - Classic Pumpkin - 550g x 4'
                                code = 'CCPUM'
                                size = '550g x 4'
                            st.write(f"Matched specific product: {product_name}")
                        # Handle Sweet Potato
                        elif 'Sweet Potato' in field_name:
                            product_name = 'CSP - Sweet Potato - 550g x 4'
                            code = 'CSP'
                            size = '550g x 4'
                            st.write(f"Matched specific product: {product_name}")
                        # Handle other products with sizes
                        else:
                            # Try to extract size from field name
                            if '350' in field_name:
                                size = '350g x 6'
                            elif '600' in field_name:
                                size = '600g x 4'
                            else:
                                size = '550g x 4'
                            
                            # Clean up field name
                            product_info = field_name
                            for pattern in order_patterns:
                                product_info = product_info.replace(pattern, '').strip()
                            
                            # Try to extract product name
                            product_name = product_info.strip()
                            code = ''
                            
                            # Look up in catalog if available
                            product_details = product_catalog.get(product_name)
                            if product_details:
                                code = product_details['Product Code']
                            
                            st.write(f"Matched product with size: {product_name}")
                            product_info = field_name
                            for pattern in order_patterns:
                                product_info = product_info.replace(pattern, '').strip()
                            
                            # Try to match additional patterns
                            if '550' in product_info:
                                if 'Pumpkin' in product_info:
                                    product_info = product_info.replace('550', '').strip()
                                    product_info = product_info.replace('x 4', '').strip()
                                    if 'Classic' in product_info:
                                        product_name = "CCPUM - Classic Pumpkin - 550g x 4"
                                        code = "CCPUM"
                                        size = "550g x 4"
                                    elif 'Sweet' in product_info:
                                        product_name = "CSP - Sweet Potato - 550g x 4"
                                        code = "CSP"
                                        size = "550g x 4"
                                else:
                                    product_name = product_info.strip()
                                    code = ""
                                    size = ""
                            elif '350' in product_info:
                                if 'Pumpkin' in product_info:
                                    product_info = product_info.replace('350', '').strip()
                                    product_info = product_info.replace('x 6', '').strip()
                                    product_name = "CPCPUM - Classic Pumpkin - 350g x 6"
                                    code = "CPCPUM"
                                    size = "350g x 6"
                                else:
                                    product_name = product_info.strip()
                                    code = ""
                                    size = ""
                            else:
                                # Default case - use the cleaned product name
                                product_name = product_info.strip()
                                code = ""
                                size = ""
                            
                            # Add to line items
                            line_items.append({
                                'Product': product_name,
                                'Quantity': quantity,
                                'Product Code': code,
                                'Size': size
                            })
            except (ValueError, TypeError):
                pass
        
        # Calculate statistics
        total_products = len(line_items)
        total_quantity = sum(item['Quantity'] for item in line_items)
        
        return {
            'customer_info': customer_info,
            'line_items': line_items,
            'stats': {
                'total_products': total_products,
                'total_quantity': total_quantity
            }
        }
        
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        return None
    finally:
        # Ensure we clean up any open streams
        if 'pdf_stream' in locals():
            pdf_stream.close()

def export_to_csv(data):
    """Export order data to CSV"""
    try:
        # Create a DataFrame from line items
        df = pd.DataFrame(data['line_items'])
        
        # Add customer information as columns
        customer_info = data['customer_info']
        for key, value in customer_info.items():
            df[key] = value
        
        # Add total statistics
        stats = data['stats']
        df['Total Products'] = stats['total_products']
        df['Total Quantity'] = stats['total_quantity']
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"order_{timestamp}.csv"
        
        # Save to CSV
        df.to_csv(filename, index=False)
        
        return filename
    except Exception as e:
        st.error(f"Error exporting to CSV: {str(e)}")
        return None

def main():
    st.title("PDF Order Form Extractor")
    
    # File upload
    uploaded_file = st.file_uploader("Upload PDF Order Form", type=['pdf'])
    
    if uploaded_file is not None:
        # Save uploaded file temporarily
        temp_path = "temp_order.pdf"
        with open(temp_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        # Extract data
        data = extract_pdf_data(temp_path)
        
        if data:
            # Display customer information
            st.subheader("Customer Information")
            customer_info = data['customer_info']
            for key, value in customer_info.items():
                st.write(f"**{key}:** {value}")
            
            # Display order statistics
            st.subheader("Order Summary")
            stats = data['stats']
            col1, col2 = st.columns(2)
            col1.metric("Total Products", stats['total_products'])
            col2.metric("Total Quantity", stats['total_quantity'])
            
            # Display line items
            st.subheader("Order Details")
            if data['line_items']:
                # Create DataFrame with proper column order
                df = pd.DataFrame(data['line_items'], columns=['Product', 'Quantity', 'Product Code', 'Size'])
                
                # Debug: Show the raw line items data
                st.write("Raw Line Items Data:", data['line_items'])
                
                # Display the DataFrame
                st.dataframe(df)
                
                # Add CSV export button
                if st.button('Export to CSV'):
                    filename = export_to_csv(data)
                    if filename:
                        st.success(f"Order exported to {filename}")
                        st.download_button(
                            label="Download CSV",
                            data=open(filename, 'rb').read(),
                            file_name=filename,
                            mime='text/csv'
                        )
            else:
                st.info("No items ordered")
        
        # Clean up
        import os
        if os.path.exists(temp_path):
            os.remove(temp_path)

if __name__ == "__main__":
    main()
