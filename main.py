# Start by importing the win32com package
import win32com.client as wincom
import requests
import json
if __name__ == "__main__":
        speak = wincom.Dispatch("SAPI.SpVoice")
        print("Welcome to Weather Forecast...")
        city = input("City Name: ")
        url=f"http://api.weatherapi.com/v1/current.json?key=b13989793f184149a91141538230103&q={city}"
        r=requests.get(url)
        wdic =json.loads(r.text)
        text=wdic["current"]["temp_c"]
        print(text)
        speak.Speak(f"temperature of {city} is {text}")
        

    