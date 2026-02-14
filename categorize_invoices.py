import pytesseract
from PIL import Image
import json
import csv
import re
import os
from datetime import datetime
from pathlib import Path

# Configure tesseract path if needed (Windows)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def find_folders_with_images(root_dir='invoices', extensions=('.jpg', '.png', '.jpeg')):
    """Find all folders containing image files recursively"""
    folders_with_images = []

    for dirpath, dirnames, filenames in os.walk(root_dir):
        # Check if current directory has image files
        has_images = any(f.lower().endswith(extensions) for f in filenames)

        if has_images:
            # Store both full path and relative path
            rel_path = os.path.relpath(dirpath, root_dir)
            folders_with_images.append({
                'full_path': dirpath,
                'relative_path': rel_path,
                'name': os.path.basename(dirpath),
                'image_count': sum(1 for f in filenames if f.lower().endswith(extensions))
            })

    return sorted(folders_with_images, key=lambda x: x['full_path'])

def display_folder_menu(folders):
    """Display interactive menu for folder selection"""
    print("\n" + "="*60)
    print(" "*18 + "FOLDER SELECTION MENU")
    print("="*60 + "\n")

    if not folders:
        print("No folders with images found!")
        return None

    for idx, folder in enumerate(folders, 1):
        print(f"  [{idx}] {folder['relative_path']}")
        print(f"      --> {folder['image_count']} images\n")

    print("="*60)

    while True:
        try:
            choice = input(f"\nSelect folder (1-{len(folders)}) or 'q' to quit: ").strip()

            if choice.lower() == 'q':
                print("Exiting...")
                return None

            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(folders):
                selected = folders[choice_idx]
                print(f"\n[OK] Selected: {selected['relative_path']}")
                print(f"     Processing {selected['image_count']} images...\n")
                return selected
            else:
                print(f"Please enter a number between 1 and {len(folders)}")
        except ValueError:
            print("Invalid input. Please enter a number or 'q' to quit.")
        except KeyboardInterrupt:
            print("\n\nOperation cancelled.")
            return None

def load_categories(config_file='categories.json'):
    """Load category keywords from config file"""
    with open(config_file, 'r') as f:
        return json.load(f)

def extract_text_from_image(image_path):
    """Extract text from image using OCR"""
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        return text.lower()
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return ""

def categorize_invoice(text, categories):
    """Categorize invoice based on keywords in text"""
    matched_categories = []

    for category_name, keywords in categories.items():
        for keyword in keywords:
            if keyword.lower() in text:
                matched_categories.append(category_name)
                break

    return matched_categories if matched_categories else ["Uncategorized"]

def filter_description(text):
    """Filter out irrelevant fields and extract description"""
    lines = text.split('\n')
    filtered_lines = []

    # Keywords to skip
    skip_patterns = [
        r'invoice\s*(no|number|#)',
        r'date',
        r'tax\s*id',
        r'iban',
        r'swift',
        r'bill\s*to',
        r'ship\s*to',
        r'payment\s*terms',
        r'due\s*date',
        r'customer\s*(no|id)',
        r'order\s*(no|number)',
        r'\d{4}-\d{2}-\d{2}',  # Date patterns
        r'^\s*$'  # Empty lines
    ]

    for line in lines:
        line_lower = line.lower().strip()
        if not line_lower:
            continue

        # Skip if line matches any irrelevant pattern
        should_skip = False
        for pattern in skip_patterns:
            if re.search(pattern, line_lower):
                should_skip = True
                break

        if not should_skip:
            filtered_lines.append(line.strip())

    return ' '.join(filtered_lines)

def save_to_csv(results, filename):
    """Save results to CSV file"""
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Filename', 'Category', 'Description'])

        for result in results:
            categories_str = ', '.join(result['categories'])
            writer.writerow([
                result['filename'],
                categories_str,
                result['description']
            ])

def process_invoices(invoice_dir='invoices', folder_name='invoices'):
    """Process all invoices and categorize them"""
    categories = load_categories()
    results = []

    # Process each invoice
    invoice_files = list(Path(invoice_dir).glob('*.jpg')) + list(Path(invoice_dir).glob('*.png'))

    if not invoice_files:
        print(f"No image files found in {invoice_dir}")
        return

    for idx, invoice_path in enumerate(invoice_files, 1):
        print(f"Processing {idx}/{len(invoice_files)}: {invoice_path.name}")

        # Extract text
        text = extract_text_from_image(invoice_path)

        # Categorize
        matched_categories = categorize_invoice(text, categories)

        # Filter description (remove invoice no, dates, tax IDs, etc.)
        description = filter_description(text)

        # Save result
        result = {
            'filename': invoice_path.name,
            'categories': matched_categories,
            'description': description,
            'full_text': text
        }
        results.append(result)

        print(f"  → {matched_categories[0]}")

    # Generate timestamp for filenames
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Clean folder name for filename (replace spaces and special chars)
    clean_folder_name = re.sub(r'[^\w\-]', '_', folder_name)

    json_filename = f'{clean_folder_name}_{timestamp}.json'
    csv_filename = f'{clean_folder_name}_{timestamp}.csv'

    # Create results directory
    Path('results').mkdir(exist_ok=True)
    json_path = Path('results', json_filename)
    csv_path = Path('results', csv_filename)

    # Save results to JSON
    with open(json_path, 'w') as f:
        json.dump(results, f, indent=2)

    # Save results to CSV
    save_to_csv(results, csv_path)

    # Print summary
    print("\n" + "="*50)
    print("CATEGORIZATION SUMMARY")
    print("="*50)
    category_counts = {}
    for result in results:
        for cat in result['categories']:
            category_counts[cat] = category_counts.get(cat, 0) + 1

    for category, count in sorted(category_counts.items()):
        print(f"{category}: {count} invoices")

    print(f"\nResults saved to:")
    print(f"  - {json_path}")
    print(f"  - {csv_path}")

if __name__ == "__main__":
    print("\n" + "="*60)
    print(" "*15 + "INVOICE CATEGORIZATION TOOL")
    print("="*60)

    # Find all folders with images
    folders = find_folders_with_images()

    if not folders:
        print("\nNo folders with images found in 'invoices' directory.")
        print("Please ensure your images are organized in folders.")
        exit(1)

    # Display menu and get user selection
    selected_folder = display_folder_menu(folders)

    if selected_folder is None:
        exit(0)

    # Process the selected folder
    process_invoices(
        invoice_dir=selected_folder['full_path'],
        folder_name=selected_folder['relative_path']
    )
