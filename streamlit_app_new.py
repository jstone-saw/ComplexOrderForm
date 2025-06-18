import streamlit as st
from pypdf import PdfReader
import pandas as pd
from io import BytesIO
import re
from datetime import datetime

# Load product catalog from Excel file
def load_product_catalog():
    try:
        catalog_df = pd.read_excel('ProductCatalog.xlsx')
        product_catalog = {}
        for _, row in catalog_df.iterrows():
            product_name = str(row['Product Name']).strip()
            if not product_name:
                continue
            product_catalog[product_name] = {
                'Product Code': str(row['Code']).strip(),
                'Size': str(row['Size']).strip()
            }
        return product_catalog
    except Exception as e:
        st.error(f"Error loading product catalog: {str(e)}")
        return {}

def extract_pdf_data(pdf_path):
    try:
        # Load product catalog
        product_catalog = load_product_catalog()
        
        # Initialize data structures
        order_data = {}
        line_items = []
        customer_info = {}
        
        # Read PDF file
        with open(pdf_path, 'rb') as file:
            pdf_data = file.read()
        pdf_stream = BytesIO(pdf_data)
        pdf_reader = PdfReader(pdf_stream)
        
        # Extract form fields
        form_fields = pdf_reader.get_fields()
        
        # Process customer information
        customer_fields = [
            'Customer Name', 'Company', 'Email', 'Phone',
            'customer_name', 'company', 'email', 'phone',
            'name', 'customer', 'contact'
        ]
        for field in customer_fields:
            if field in form_fields:
                customer_info[field] = form_fields[field].get('/V', '')
        
        # Define product patterns
        product_patterns = {
            # Soup products
            'Ancient Grain': ['Ancient Grain', 'CAGR', 'agr', 'ancient'],
            'Beef & Barley': ['Beef & Barley', 'CBAB', 'bab', 'beef'],
            'Chicken & Corn': ['Chicken & Corn', 'CCCORN', 'corn', 'chicken'],
            'Chicken Noodle': ['Chicken Noodle', 'CNCH', 'noodle', 'chickennoodle'],
            'Chunky Minestrone': ['Chunky Minestrone', 'CCMIN', 'minestrone', 'min'],
            'Creamy Chicken': ['Creamy Chicken', 'CCRC', 'creamychicken', 'cc'],
            'Creamy Mushroom': ['Creamy Mushroom', 'CMUSH', 'mushroom', 'cm'],
            'Creamy Pumpkin': ['Creamy Pumpkin', 'CRMPUM', 'pumpkin', 'crmp'],
            'French Onion': ['French Onion', 'CFON', 'onion', 'fon'],
            'Hearty Beef': ['Hearty Beef', 'CHB', 'heartybeef', 'hb'],
            # Add more patterns as needed
        }

        # Process each field
        for field_name, field in form_fields.items():
            try:
                value = field.get('/V', '')
                if not value:
                    continue

                # Try to parse quantity
                quantity = None
                try:
                    quantity = int(value)
                except ValueError:
                    # Try to extract numbers from string
                    numbers = re.findall(r'\d+', str(value))
                    if numbers:
                        quantity = int(numbers[0])
                
                if quantity is None or quantity <= 0:
                    continue

                st.write(f"Processing field: {field_name} with quantity: {quantity}")

                # Try to match field name with product patterns
                matched_product = None
                for product_name, patterns in product_patterns.items():
                    if any(pattern.lower() in field_name.lower() for pattern in patterns):
                        matched_product = product_name
                        break

                if matched_product:
                    # Look up product details from catalog
                    product_info = product_catalog.get(matched_product)
                    if product_info:
                        line_items.append({
                            'Product Name': matched_product,
                            'Product Code': product_info['Product Code'],
                            'Size': product_info['Size'],
                            'Quantity': quantity
                        })
                    else:
                        st.warning(f"Product '{matched_product}' found in PDF but not in catalog")
                else:
                    st.warning(f"Could not match field '{field_name}' to any known product")

            except Exception as e:
                st.error(f"Error processing field {field_name}: {str(e)}")
                continue

        # Calculate statistics
        total_products = len(line_items)
        total_quantity = sum(item['Quantity'] for item in line_items)

        # Store processed data
        order_data.update({
            'customer_info': customer_info,
            'line_items': line_items,
            'stats': {
                'total_products': total_products,
                'total_quantity': total_quantity
            }
        })

        return order_data

    except Exception as e:
        st.error(f"Error extracting PDF data: {str(e)}")
        return None

# Export order data to CSV
def export_to_csv(data):
    try:
        df = pd.DataFrame(data['line_items'])
        
        # Add customer information as columns
        customer_info = data['customer_info']
        for key, value in customer_info.items():
            df[key] = value
        
        # Add statistics
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
                df = pd.DataFrame(data['line_items'])
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
