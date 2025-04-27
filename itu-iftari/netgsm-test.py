#!/usr/bin/env python
"""
Bulk SMS sending example for Netgsm Python SDK.
"""

import os

from dotenv import load_dotenv
from netgsm import Netgsm

# Load environment variables
load_dotenv()

HEADER = f"0{os.getenv('NETGSM_USERNAME')}"


def main():
    """Main function."""
    # Initialize Netgsm client
    client = Netgsm(
        username=os.getenv("NETGSM_USERNAME"),
        password=os.getenv("NETGSM_USER_PASSWORD"),
        # header=HEADER
    )

    # Example phone numbers (replace with actual numbers)
    phone_numbers = [
        "5313704175",
        "+90 534 518 81 92",
    ]

    # Example messages for each phone number
    messages = [
        {
            "msg": f"Test message {i + 1} from Netgsm Python SDK",
            "no": phone,
        }
        for i, phone in enumerate(phone_numbers)
    ]

    try:
        # Send bulk SMS
        response = client.sms.send(
            messages=messages,
            msgheader=HEADER,
        )
        print("Bulk SMS sent successfully!")
        print("Response:", response)
    except Exception as e:
        print("Error sending bulk SMS:", e)


if __name__ == "__main__":
    main()
