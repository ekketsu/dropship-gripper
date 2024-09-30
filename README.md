# Ship-Grip

**ship-grip** is an automated tool designed to collect product information from e-commerce website saves it to an Excel file.
## Features

- Collects product information from website, including:
  - Product name
  - Price
  - Checks for duplicates with an existing Excel file and only collects new products.
  - Downloads product images and saves them in a subfolder.
  - Uses Selenium to automate page scrolling and product collection.
    
  ### Coming Soon

  - Number of sales
  - Ratings
  - Number of customer reviews
  - Seller name
  - Shipping location

## Prerequisites

- Python 3.x
- Google Chrome (for Selenium)
- ChromeDriver (automatically installed with `webdriver-manager`)

## Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/ekketsu/ship-grip.git
   ```

2. Navigate to the project directory:
   ```bash
   cd ship-grip
   ```

3. Install the dependencies using `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

To run the script, use the following command:
```bash
python3 ship-grip.py
```

This will collect up to 60 new products from the specified URL and save the results in the file `products.xlsx`.

## Command-Line Options

- `--max-products` : Maximum number of new products to collect. *(default: 60)*
- `--url` : URL of the AliExpress page to scrape. *(default: `https://www.aliexpress.com/p/calp-plus/index.html`)*
- `--output` : Path to the output Excel file. *(default: `products.xlsx`)*
- `--scroll-pause` : Wait time (in seconds) after each scroll to ensure product loading. *(default: 2.0)*

Example:
```bash
python3 ship-grip.py --max-products 100 --url "https://www.aliexpress.com/category/100003109/women-clothing.html" --output "new_file.xlsx" --scroll-pause 1.5
```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.
