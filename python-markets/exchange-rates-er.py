import os
import requests
from dotenv import load_dotenv

def get_currency_rate():
    print("\n==============================")
    print("🌍 Exchange Rate Checker")
    print("==============================\n")

    'https://app.exchangerate-api.com/dashboard
    load_dotenv()
    API_KEY = os.getenv("ER_API_KEY")

    if not API_KEY:
        print("⚠️  No API key found in .env (ER_API_KEY)")
        API_KEY = input("Please enter your ExchangeRate-API key: ").strip()

    while True:
        base_currency = input("Enter base currency (e.g., EUR, USD), or press Enter to exit: ").upper().strip()

        if base_currency == "":
            print("\n👋 Exiting. Have a great day!")
            break

        target_currency = input("Enter target currency (e.g., PLN): ").upper().strip()

        url = f"https://v6.exchangerate-api.com/v6/{API_KEY}/latest/{base_currency}"

        try:
            response = requests.get(url)
            data = response.json()

            if data.get("result") == "success":
                rate = data["conversion_rates"].get(target_currency)
                if rate:
                    print("\n🌍 Exchange Rate")
                    print("============================")
                    print(f"💱 1 {base_currency} = {rate:.4f} {target_currency}")
                    print("============================\n")
                else:
                    print(f"❌ Currency '{target_currency}' not found.")
            else:
                print(f"❌ API Error: {data.get('error-type', 'Unknown error')}")

        except Exception as e:
            print(f"❌ Request failed: {e}")

if __name__ == "__main__":
    get_currency_rate()
