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
                        
                        # Check if this product already exists in line_items
                        existing_item = None
                        for item in line_items:
                            if (item['Product Name'] == matched_product and 
                                item['Size'] == size and 
                                item['Product Code'] == product_info['Product Code']):
                                existing_item = item
                                break
                        
                        if existing_item:
                            # If it exists, add the quantity to the existing item
                            existing_item['Quantity'] += quantity
                        else:
                            # If it doesn't exist, add a new item
                            line_items.append({
                                'Product Name': matched_product,
                                'Product Code': product_info['Product Code'],
                                'Size': size,
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
