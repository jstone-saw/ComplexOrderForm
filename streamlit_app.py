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
        
        # Create mapping of all known products from PDF
        known_products = {
            'Ancient Grain', 'Beef & Barley', 'Chicken & Corn', 'Chicken Noodle',
            'Chunky Minestrone', 'Creamy Chicken', 'Creamy Mushroom', 'Creamy Pumpkin',
            'French Onion', 'Hearty Beef', 'Hearty Chicken', 'Hearty Vegetable',
            'Lentil', 'Pea & Ham', 'Potato & Leek', 'Sweet Potato', 'Tomato',
            'Chixken & Vegetable 350', 'Classic Pumpkin', 'Tomato 350',
            'Asian Chicken', 'Crab Chowder', 'Chicken Laksa', 'Thai Prawn',
            'Asian Chicken 600', 'Crab & Corn 600', 'Thai Prawn 600',
            'Moroccan Harira 600', 'Chicken Tagine', 'Red Chicken Curry',
            'Red Vegetable Curry', 'Hummous 350', 'Chilli Lemon Hummous 330',
            'Chunky Eggplant Hummous 330', 'Olive Salsa Hummous 330',
            'Garlic Dip 310', 'Taramosalata 310', 'Beetroot Almond 200',
            'Chunky Hummous 200', 'Capsicum Salsa 200', 'Eggplant Capsicum 200',
            'Eggplant Hummous 200', 'Hummous 200', 'Harissa Hummous 200',
            'Pine Nut Hummous 200', 'Spicy Carrot 200', 'Garlic Dip 180',
            'Taramosalata 180', 'Smoked Salmon 170', 'Baba Ganoush',
            'Beetroot Hummus', 'Chilli Kalamata Hummus', 'Harissa Hummus',
            'Mediterranean Hummus', 'Olive Hummus', 'Pine Nut Hummus',
            'Roasted Garlic Hummus', 'Hummus 1kg'
        }
        
        # Load products from Excel
        for _, row in catalog_df.iterrows():
            product_name = str(row['Product Name']).strip()
            if not product_name:
                continue
                
            # Add product to catalog if it's one of our known products
            if product_name in known_products:
                product_catalog[product_name] = {
                    'Product Code': str(row['Code']).strip(),
                    'Size': str(row['Size']).strip()
                }
                
                # Also add variations of the product name
                variations = {
                    # Handle common misspellings
                    'Chixken': 'Chicken',
                    'Hummous': 'Hummus'
                }
                
                for misspelling, correct in variations.items():
                    if misspelling in product_name:
                        corrected_name = product_name.replace(misspelling, correct)
                        product_catalog[corrected_name] = product_catalog[product_name]
        
        # Add any missing products with default values
        for product in known_products:
            if product not in product_catalog:
                product_catalog[product] = {
                    'Product Code': f'C{product[:4].upper()}',  # Generate a default code
                    'Size': 'Unknown'  # Will be updated later
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
        order_data = []  # Changed to a list of dictionaries
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
            'name', 'customer', 'contact',
            'Date', 'date', 'Order Date', 'order_date'
        ]
        
        # Process each field to extract customer info
        for field_name, field in form_fields.items():
            value = field.get('/V', '')
            if not value:
                continue
                
            # Try to match with customer fields
            for field in customer_fields:
                if field.lower() in field_name.lower():
                    # Special handling for date field
                    if 'date' in field.lower():
                        try:
                            # Try to parse date in format DD/MM/YYYY
                            date_value = value.strip()
                            if '/' in date_value:
                                customer_info['Order Date'] = date_value
                                break
                        except:
                            continue
                    else:
                        customer_info[field] = value
                        break
        
        # Define comprehensive product patterns
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
            'Hearty Chicken': ['Hearty Chicken', 'CHCK', 'chicken', 'chixken'],
            'Hearty Vegetable': ['Hearty Vegetable', 'CHVEG', 'vegetable'],
            'Lentil': ['Lentil', 'CLEN', 'lentils'],
            'Pea & Ham': ['Pea & Ham', 'CPH', 'pea', 'ham'],
            'Potato & Leek': ['Potato & Leek', 'CPLE', 'potato', 'leek'],
            'Sweet Potato': ['Sweet Potato', 'CSP', 'sweet'],
            'Tomato': ['Tomato', 'CTOM', 'tomatoes'],
            'Asian Chicken': ['Asian Chicken', 'CACH', 'asian'],
            'Crab Chowder': ['Crab Chowder', 'CCCH', 'crab'],
            'Chicken Laksa': ['Chicken Laksa', 'CLAK', 'laksa'],
            'Thai Prawn': ['Thai Prawn', 'CTHP', 'prawn', 'thai'],
            'Chicken & Vegetable 350': ['Chicken & Vegetable 350', 'CCHV350', 'chickenveg350'],
            'Tomato 350': ['Tomato 350', 'CTOM350', 'tomato350'],
            'Asian Chicken 600': ['Asian Chicken 600', 'CACH600', 'asian600'],
            'Crab & Corn 600': ['Crab & Corn 600', 'CCCH600', 'crabcorn600'],
            'Thai Prawn 600': ['Thai Prawn 600', 'CTHP600', 'prawn600'],
            'Moroccan Harira 600': ['Moroccan Harira 600', 'CMOR600', 'moroccan600'],
            'Chicken Tagine': ['Chicken Tagine', 'CCTG', 'tagine'],
            'Red Chicken Curry': ['Red Chicken Curry', 'CRCC', 'redchicken'],
            'Red Vegetable Curry': ['Red Vegetable Curry', 'CRVC', 'redveg'],
            'Hummous 350': ['Hummous 350', 'CHM350', 'hummous350'],
            'Chilli Lemon Hummous 330': ['Chilli Lemon Hummous 330', 'CHLH330', 'chillilemon330'],
            'Chunky Eggplant Hummous 330': ['Chunky Eggplant Hummous 330', 'CEHH330', 'chunkyeggplant330'],
            'Olive Salsa Hummous 330': ['Olive Salsa Hummous 330', 'COHH330', 'olivesalsa330'],
            'Garlic Dip 310': ['Garlic Dip 310', 'CGD310', 'garlicdip310'],
            'Taramosalata 310': ['Taramosalata 310', 'CTA310', 'taramosalata310'],
            'Beetroot Almond 200': ['Beetroot Almond 200', 'CBRA200', 'beetroot200'],
            'Chunky Hummous 200': ['Chunky Hummous 200', 'CHH200', 'chunky200'],
            'Capsicum Salsa 200': ['Capsicum Salsa 200', 'CCS200', 'capsicum200'],
            'Eggplant Capsicum 200': ['Eggplant Capsicum 200', 'CEC200', 'eggplantcapsicum200'],
            'Eggplant Hummous 200': ['Eggplant Hummous 200', 'CEH200', 'eggplanthummous200'],
            'Hummous 200': ['Hummous 200', 'CHM200', 'hummous200'],
            'Harissa Hummous 200': ['Harissa Hummous 200', 'CHH200', 'harissa200'],
            'Pine Nut Hummous 200': ['Pine Nut Hummous 200', 'CPH200', 'pinenut200'],
            'Spicy Carrot 200': ['Spicy Carrot 200', 'CSC200', 'spicycarrot200'],
            'Garlic Dip 180': ['Garlic Dip 180', 'CGD180', 'garlicdip180'],
            'Taramosalata 180': ['Taramosalata 180', 'CTA180', 'taramosalata180'],
            'Smoked Salmon 170': ['Smoked Salmon 170', 'CSS170', 'smokedsalmon170'],
            'Baba Ganoush': ['Baba Ganoush', 'CBG', 'baba'],
            'Beetroot Hummus': ['Beetroot Hummus', 'CBRH', 'beetroot'],
            'Chilli Kalamata Hummus': ['Chilli Kalamata Hummus', 'CCHK', 'chillikalamata'],
            'Harissa Hummus': ['Harissa Hummus', 'CHH', 'harissa'],
            'Mediterranean Hummus': ['Mediterranean Hummus', 'CMH', 'mediterranean'],
            'Olive Hummus': ['Olive Hummus', 'COH', 'olive'],
            'Pine Nut Hummus': ['Pine Nut Hummus', 'CPH', 'pinenut'],
            'Roasted Garlic Hummus': ['Roasted Garlic Hummus', 'CRGH', 'roastedgarlic'],
            'Hummus 1kg': ['Hummus 1kg', 'CHM1KG', 'hummus1kg']
        }

        # Track processed fields and products
        processed_fields = set()
        processed_products = {}
        
        # Process each field
        for field_name, field in form_fields.items():
            try:
                value = field.get('/V', '')
                if not value:
                    continue

                # Try to parse quantity
                quantity = None
                try:
                    # Try multiple quantity parsing methods
                    # 1. Direct integer conversion
                    quantity = int(value)
                except ValueError:
                    try:
                        # 2. Extract numbers from string
                        numbers = re.findall(r'\d+', str(value))
                        if numbers:
                            quantity = int(numbers[0])
                    except:
                        try:
                            # 3. Handle decimal numbers
                            quantity = int(float(value))
                        except:
                            pass
                
                if quantity is None or quantity <= 0:
                    continue

                # Try to match field name with product patterns
                matched_product = None
                matched_pattern = None
                
                # Find the most specific pattern match
                for product_name, patterns in product_patterns.items():
                    for pattern in patterns:
                        if pattern.lower() in field_name.lower():
                            # Check if this is a more specific match than previous matches
                            if matched_product is None or len(pattern) > len(matched_pattern):
                                matched_product = product_name
                                matched_pattern = pattern
                                break

                if matched_product:
                    # Look up product details from catalog
                    product_info = product_catalog.get(matched_product)
                    if product_info:
                        # Get size from field name if it contains a size specification
                        size = product_info['Size']
                        if '350' in field_name.lower():
                            size = "350g"
                        elif '600' in field_name.lower():
                            size = "600g"
                        elif '200' in field_name.lower():
                            size = "200g"
                        elif '180' in field_name.lower():
                            size = "180g"
                        elif '170' in field_name.lower():
                            size = "170g"
                        elif '1kg' in field_name.lower():
                            size = "1kg"
                        
                        # Create unique key for this product variant
                        product_key = f"{field_name}_{matched_product}_{size}"
                        
                        # Check if we've already processed this exact field
                        if field_name in processed_fields:
                            continue
                        
                        # Create new order item
                        order_item = {
                            'Product Name': matched_product,
                            'Product Code': product_info['Product Code'],
                            'Size': size,
                            'Quantity': quantity,
                            **customer_info  # Include all customer info
                        }
                        order_data.append(order_item)
                        processed_fields.add(field_name)
                        processed_products[product_key] = order_item
                    else:
                        continue
                else:
                    continue

            except Exception as e:
                st.error(f"Error processing field {field_name}: {str(e)}")
                continue

        # Calculate statistics
        total_products = len(order_data)
        total_quantity = sum(item['Quantity'] for item in order_data)

        # Add statistics as a separate item
        stats_item = {
            'Product Name': 'Total',
            'Product Code': '',
            'Size': '',
            'Quantity': total_quantity,
            'Total Products': total_products,
            **customer_info
        }
        order_data.append(stats_item)

        return order_data

    except Exception as e:
        st.error(f"Error extracting PDF data: {str(e)}")
        return []
        return None

# Export order data to CSV
def export_to_csv(data):
    try:
        if not data:
            st.error("No data to export")
            return None
            
        # Create DataFrame directly from the flattened data
        df = pd.DataFrame(data)
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"order_{timestamp}.csv"
        
        # Save to CSV with proper formatting
        df.to_csv(filename, index=False, encoding='utf-8')
        
        return filename
    except Exception as e:
        st.error(f"Error exporting to CSV: {str(e)}")
        return None
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
            # Extract customer info from first item since it's consistent across all items
            customer_info = data[0]
            for key in ['Customer Name', 'Company', 'Email', 'Phone', 'Order Date']:
                if key in customer_info:
                    st.write(f"**{key}:** {customer_info[key]}")
            
            # Display order statistics
            st.subheader("Order Summary")
            # Find the stats item (it has 'Total' as Product Name)
            stats_item = next((item for item in data if item.get('Product Name') == 'Total'), None)
            if stats_item:
                col1, col2 = st.columns(2)
                col1.metric("Total Products", stats_item.get('Total Products', 0))
                col2.metric("Total Quantity", stats_item.get('Quantity', 0))
            
            # Display line items
            st.subheader("Order Details")
            # Create DataFrame from all items except the stats item
            df = pd.DataFrame([item for item in data if item.get('Product Name') != 'Total'])
            
            # Reorder columns
            column_order = ['Customer Name', 'Order Date', 'Product Code', 'Product Name', 'Size', 'Quantity']
            df = df[column_order]
            
            st.dataframe(df)
            
            # Add debugging info under the table
            st.subheader("Extracted Items")
            for item in data:
                if item.get('Product Name') != 'Total':
                    st.write(f"Field: {item.get('Product Name')} ({item.get('Size')}) - Quantity: {item.get('Quantity')}")
            
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
