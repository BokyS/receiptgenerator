import json
import base64
import urllib.request


def generate(c_data):
    api_url = "https://hub3.bigfish.software/api/v2/barcode?data="
    f = {
        "renderer": "image",
        "options": {
            "format": "png",
            "scale": 3,
            "ratio": 3,
            "color": "#2c3e50",
            "bgColor": "#eeeeee",
            "padding": 20
        },
        "data": {
            "amount": int(c_data["amount_tax"]*100),
            "currency": c_data["currency"],
            "sender": {
                "name": c_data["name"],
                "street": c_data["address"],
                "place": c_data["postalcode"]
            },
            "receiver": {
                "name": "Your company name",
                "street": "Company Adress",
                "place": "10000, Zagreb", #Postal code, City
                "iban": "HR0000000000000000000", #Company IBAN
                "model": "00", #Payment model
                "reference": c_data["refrence"] #Refrance to receipt
            },
            "purpose": "SCVE", #Purchase & Sale of Services code
            "description": c_data["description"] #Additional description
        }
    }
    pack = json.dumps(f)
    req = base64.b64encode(pack.encode('utf-8'))
    whole_url = api_url + (req.decode('ascii'))
    urllib.request.urlretrieve(whole_url, "barcode\\" + c_data["refrence"] + c_data["name"].replace(" ","") + ".jpg")

