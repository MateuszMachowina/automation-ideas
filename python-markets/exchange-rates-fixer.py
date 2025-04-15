import os
import requests
from dotenv import load_dotenv

def get_currency_rate():
    print("\n==============================")
    print("🌍 EUR Exchange Rate Checker")
    print("==============================\n")

    load_dotenv()
    API_KEY = os.getenv("FIXER_API_KEY")

    if not API_KEY:
        print("⚠️  No API key found in .env (FIXER_API_KEY)")
        API_KEY = input("Please enter your Fixer.io API key: ").strip()

    base_currency = "EUR"

    while True:
        target_currency = input("Enter target currency (e.g., PLN), or press Enter to exit: ").upper().strip()

        if target_currency == "":
            print("\n👋 Exiting. Have a great day!")
            break

        url = f"http://data.fixer.io/api/latest?access_key={API_KEY}&symbols={target_currency}"

        try:
            response = requests.get(url)
            data = response.json()

            if data.get("success"):
                rate = data["rates"].get(target_currency)
                if rate:
                    print("\n🌍 Exchange Rate")
                    print("============================")
                    print(f"💱 1 EUR = {rate:.4f} {target_currency}")
                    print("============================\n")
                else:
                    print(f"❌ Currency '{target_currency}' not found.")
            else:
                print(f"❌ API Error: {data.get('error', {}).get('info', 'Unknown error')}")

        except Exception as e:
            print(f"❌ Request failed: {e}")

if __name__ == "__main__":
    get_currency_rate()
