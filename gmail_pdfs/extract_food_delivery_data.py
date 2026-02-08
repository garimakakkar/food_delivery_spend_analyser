#!/usr/bin/env python3
"""
Extract data from Zomato invoice PDFs and export to Google Sheets
"""

import pdfplumber
import os
import re
from datetime import datetime
from typing import List, Dict, Optional
import csv

def extract_order_data(pdf_path: str) -> Optional[Dict]:
    """Extract order data from a Zomato PDF invoice"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            
            if not text:
                return None
            
            data = {
                'order_id': None,
                'date': None,
                'restaurant_name': None,
                'items': [],
                'total': None,
                'file_name': os.path.basename(pdf_path)
            }
            
            # Extract Order ID
            order_id_match = re.search(r'Order ID:\s*(\d+)', text)
            if order_id_match:
                data['order_id'] = order_id_match.group(1)
            
            # Extract Date/Time
            date_match = re.search(r'Order Time:\s*(\d+\s+\w+\s+\d{4}[,\s]+\d+:\d+\s*(?:AM|PM))', text)
            if date_match:
                date_str = date_match.group(1)
                try:
                    # Parse date like "31 January 2026, 12:52 PM"
                    date_str_clean = date_str.replace(',', '')
                    data['date'] = datetime.strptime(date_str_clean, "%d %B %Y %I:%M %p").strftime("%Y-%m-%d")
                except:
                    # Try alternative format
                    try:
                        data['date'] = datetime.strptime(date_str_clean, "%d %b %Y %I:%M %p").strftime("%Y-%m-%d")
                    except:
                        data['date'] = date_str
            
            # Extract Restaurant Name
            restaurant_match = re.search(r'Restaurant Name:\s*([^\n]+)', text)
            if restaurant_match:
                data['restaurant_name'] = restaurant_match.group(1).strip()
            
            # Extract Items - look for item lines
            # Pattern: Item name, quantity, unit price, total price
            lines = text.split('\n')
            items_section = False
            items_list = []
            
            for i, line in enumerate(lines):
                # Detect start of items section
                if 'Item' in line and 'Quantity' in line and 'Price' in line:
                    items_section = True
                    continue
                
                # Stop at charges section (but not "Total Price" header)
                if items_section:
                    # Stop at charges, but skip "Total Price" header
                    if 'total price' in line.lower() and ('quantity' in line.lower() or 'unit' in line.lower()):
                        continue  # This is the header, not the final total
                    if any(keyword in line.lower() for keyword in ['taxes', 'packaging', 'delivery', 'platform', 'round off']):
                        items_section = False
                        continue
                    if 'terms' in line.lower() or 'conditions' in line.lower():
                        items_section = False
                        continue
                    
                    # Extract item line
                    if items_section and line.strip():
                        # Look for pattern: Item name followed by numbers
                        item_match = re.match(r'^([^₹]+?)\s+(\d+)\s+₹\s*([\d,]+\.?\d*)\s+₹\s*([\d,]+\.?\d*)', line)
                        if item_match:
                            item_name = item_match.group(1).strip()
                            quantity = item_match.group(2)
                            unit_price = item_match.group(3).replace(',', '')
                            total_price = item_match.group(4).replace(',', '')
                            items_list.append(f"{item_name} (Qty: {quantity}, ₹{total_price})")
            
            # If items weren't found in structured format, try alternative extraction
            if not items_list:
                # Look for lines with ₹ symbol that might be items
                for line in lines:
                    if '₹' in line and not any(keyword in line.lower() for keyword in ['tax', 'delivery', 'platform', 'packaging', 'total', 'free']):
                        # Try to extract item name and price
                        item_parts = re.split(r'₹', line)
                        if len(item_parts) >= 2:
                            item_name = item_parts[0].strip()
                            if item_name and len(item_name) > 3:  # Reasonable item name length
                                price_match = re.search(r'₹\s*([\d,]+\.?\d*)', line)
                                if price_match:
                                    items_list.append(f"{item_name} (₹{price_match.group(1)})")
            
            data['items'] = items_list
            
            # Find the final TOTAL - search from bottom up to get the last one
            # (which should be after all charges)
            total_patterns = [
                r'^Total\s+₹\s*([\d,]+\.?\d*)',  # Line starting with "Total ₹"
                r'^Total\s*:\s*₹\s*([\d,]+\.?\d*)',  # Line starting with "Total: ₹"
                r'^Total\s*₹\s*([\d,]+\.?\d*)',  # Line starting with "Total₹"
            ]
            
            # Search from bottom up to find the last TOTAL (which is the final total)
            for line in reversed(lines):
                line_clean = line.strip()
                if not line_clean:
                    continue
                    
                for pattern in total_patterns:
                    total_match = re.match(pattern, line_clean, re.IGNORECASE)
                    if total_match:
                        total_str = total_match.group(1).replace(',', '')
                        try:
                            total_value = float(total_str)
                            # Make sure it's a reasonable total (not too small, likely > 20)
                            if total_value > 20:
                                data['total'] = total_value
                                break
                        except:
                            continue
                if data['total']:
                    break
            
            # If still not found, try broader search
            if not data['total']:
                for pattern in total_patterns:
                    total_match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                    if total_match:
                        total_str = total_match.group(1).replace(',', '')
                        try:
                            total_value = float(total_str)
                            if total_value > 20:
                                data['total'] = total_value
                                break
                        except:
                            continue
            
            # Only return data if we have essential fields
            if data['date'] and data['restaurant_name'] and data['total']:
                return data
            
            return None
            
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return None

def process_all_pdfs(folder_path: str) -> List[Dict]:
    """Process all PDF files in the folder"""
    all_data = []
    pdf_files = []
    
    # Get all PDF files, prioritizing Order_ID and Order_Invoice files
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            # Skip User_Charge_Invoice files as they don't contain order totals
            if 'User_Charge_Invoice' not in filename:
                pdf_files.append(os.path.join(folder_path, filename))
    
    # Sort by filename to process in order
    pdf_files.sort()
    
    print(f"Found {len(pdf_files)} PDF files to process...")
    
    for pdf_path in pdf_files:
        print(f"Processing: {os.path.basename(pdf_path)}")
        data = extract_order_data(pdf_path)
        if data:
            all_data.append(data)
        else:
            print(f"  ⚠️  Could not extract data from {os.path.basename(pdf_path)}")
    
    return all_data

def export_to_csv(data: List[Dict], output_path: str):
    """Export data to CSV file"""
    if not data:
        print("No data to export!")
        return
    
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Date', 'Restaurant Name', 'Items Ordered', 'Total (₹)', 'Order ID', 'File Name']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for row in data:
            writer.writerow({
                'Date': row['date'],
                'Restaurant Name': row['restaurant_name'],
                'Items Ordered': '; '.join(row['items']) if row['items'] else 'N/A',
                'Total (₹)': row['total'],
                'Order ID': row['order_id'] or 'N/A',
                'File Name': row['file_name']
            })
    
    print(f"\n✅ Data exported to {output_path}")

def analyze_data(data: List[Dict]):
    """Analyze the extracted data"""
    if not data:
        print("No data to analyze!")
        return
    
    print("\n" + "="*80)
    print("COST ANALYSIS")
    print("="*80)
    
    # Total spending
    total_spending = sum(row['total'] for row in data)
    print(f"\n📊 Total Spending: ₹{total_spending:,.2f}")
    
    # Number of orders
    num_orders = len(data)
    print(f"📦 Number of Orders: {num_orders}")
    
    # Average order value
    avg_order = total_spending / num_orders if num_orders > 0 else 0
    print(f"💰 Average Order Value: ₹{avg_order:,.2f}")
    
    # Highest and lowest orders
    totals = [row['total'] for row in data]
    max_order = max(totals)
    min_order = min(totals)
    max_order_data = next(row for row in data if row['total'] == max_order)
    min_order_data = next(row for row in data if row['total'] == min_order)
    
    print(f"\n📈 Highest Order: ₹{max_order:,.2f}")
    print(f"   Restaurant: {max_order_data['restaurant_name']}")
    print(f"   Date: {max_order_data['date']}")
    
    print(f"\n📉 Lowest Order: ₹{min_order:,.2f}")
    print(f"   Restaurant: {min_order_data['restaurant_name']}")
    print(f"   Date: {min_order_data['date']}")
    
    # Restaurant frequency
    restaurant_counts = {}
    restaurant_totals = {}
    for row in data:
        restaurant = row['restaurant_name']
        restaurant_counts[restaurant] = restaurant_counts.get(restaurant, 0) + 1
        restaurant_totals[restaurant] = restaurant_totals.get(restaurant, 0) + row['total']
    
    print(f"\n🍽️  Top 10 Restaurants by Order Count:")
    sorted_restaurants = sorted(restaurant_counts.items(), key=lambda x: x[1], reverse=True)
    for i, (restaurant, count) in enumerate(sorted_restaurants[:10], 1):
        total_spent = restaurant_totals[restaurant]
        print(f"   {i}. {restaurant}: {count} orders, ₹{total_spent:,.2f} total")
    
    print(f"\n💵 Top 10 Restaurants by Total Spending:")
    sorted_by_spending = sorted(restaurant_totals.items(), key=lambda x: x[1], reverse=True)
    for i, (restaurant, total) in enumerate(sorted_by_spending[:10], 1):
        count = restaurant_counts[restaurant]
        print(f"   {i}. {restaurant}: ₹{total:,.2f} ({count} orders)")
    
    # Monthly breakdown
    monthly_totals = {}
    monthly_counts = {}
    for row in data:
        if row['date']:
            month = row['date'][:7]  # YYYY-MM
            monthly_totals[month] = monthly_totals.get(month, 0) + row['total']
            monthly_counts[month] = monthly_counts.get(month, 0) + 1
    
    print(f"\n📅 Monthly Breakdown:")
    for month in sorted(monthly_totals.keys()):
        total = monthly_totals[month]
        count = monthly_counts[month]
        avg = total / count if count > 0 else 0
        print(f"   {month}: {count} orders, ₹{total:,.2f} total, ₹{avg:,.2f} avg")
    
    print("\n" + "="*80)

if __name__ == "__main__":
    folder_path = "/Users/garima/Desktop/zomato/gmail_pdfs"
    
    print("Starting PDF extraction...")
    all_data = process_all_pdfs(folder_path)
    
    if all_data:
        # Sort by date
        all_data.sort(key=lambda x: x['date'] or '')
        
        # Export to CSV
        csv_path = os.path.join(folder_path, "zomato_orders.csv")
        export_to_csv(all_data, csv_path)
        
        # Analyze data
        analyze_data(all_data)
        
        print(f"\n✅ Processed {len(all_data)} orders successfully!")
    else:
        print("❌ No data extracted from PDFs!")
