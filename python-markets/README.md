
# Python Markets Scripts

For now this repository contains two Python scripts to check exchange rates from different APIs. Both scripts allow users to get the exchange rate between EUR and any other target currency.

## Scripts

1. **`exchange-rates-fixer.py`**  
   Uses the **Fixer.io API** to check exchange rates.  
   **Note**: This script only supports EUR as the base currency, and it allows you to check the rate of EUR against any target currency.

2. **`exchange-rates-er.py`**  
   Uses the **ExchangeRate-API** to check exchange rates.  
   **Note**: This script allows you to check exchange rates for various base currencies (e.g., EUR, USD, etc.) against a target currency.

## Prerequisites

Before using these scripts, you need to have the following installed:

- Python 3.13

You will also need an API key for both Fixer.io and ExchangeRate-API.

## Setup

### 1. Install Required Libraries

Clone or download this repository, and then install the required libraries using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

The `requirements.txt` file contains:

```txt
requests
python-dotenv
```

### 2. Create a `.env` File

You will need to create a `.env` file in the project directory to store your API keys securely.

#### Example of `.env` file:

```env
FIXER_API_KEY=your_fixer_api_key_here
ER_API_KEY=your_exchangerate_api_key_here
```

### 3. Running the Scripts

To run the scripts, open a terminal or command prompt in the project directory and execute one of the following commands:

- To run the `exchange-rates-fixer.py` script:

    ```bash
    python exchange-rates-fixer.py
    ```

- To run the `exchange-rates-er.py` script:

    ```bash
    python exchange-rates-er.py
    ```

### 4. Using the Scripts

- Once the script is running, you will be prompted to enter a **target currency** (e.g., PLN).
- The script will then display the exchange rate between EUR and the target currency.
- You can continue checking rates or exit the script by pressing Enter without input.

#### Example usage:

```bash
Enter target currency (e.g., PLN), or press Enter to exit: USD
🌍 Exchange Rate
============================
💱 1 EUR = 1.2000 USD
============================

Enter target currency (e.g., PLN), or press Enter to exit: GBP
🌍 Exchange Rate
============================
💱 1 EUR = 0.9000 GBP
============================

Enter target currency (e.g., PLN), or press Enter to exit:
👋 Exiting. Have a great day!
```

### 5. API Key Notes

- **Fixer.io API**: [Sign up for a Fixer.io API key](https://fixer.io/signup).
- **ExchangeRate-API**: [Sign up for a free ExchangeRate-API key](https://www.exchangerate-api.com/).

Ensure that you add the API keys to your `.env` file as shown in the setup section.
