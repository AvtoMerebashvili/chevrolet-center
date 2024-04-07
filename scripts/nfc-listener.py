from smartcard.System import readers
from smartcard.util import toHexString
import time
import xlsOpenerWithFilter as xlsHandler


def card_detected(card_data):
    card_id = str(card_data)
    print("Card detected:", toHexString(card_data))
    xlsHandler.openExcelAndFilter(card_id)


def main():
    r = readers()
    reader = r[0]
    while True:
        connection = reader.createConnection()
        try:
            connection.connect()
            data, w1,w2 = connection.transmit([0xFF, 0xCA, 0x00, 0x00, 0x00])
            card_detected(data)
        
        except Exception as e:
            # Handle exceptions (e.g., no card found)
            pass
        time.sleep(1)  # Pause for a short period to avoid rapid polling

if __name__ == "__main__":
    main()


