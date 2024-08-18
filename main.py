import pandas as pd
import requests
import time
import os

API_KEY = os.getenv("API_KEY")


def check_place(nombre):
    try:
        url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
        params = {"query": f"{nombre} Costa Rica", "key": API_KEY}
        response = requests.get(url, params=params)
        result = response.json()

        if "results" in result and len(result["results"]) > 0:
            # Извлечение первой найденной записи
            place = result["results"][0]
            status = "match"

            # Дополнительные данные
            name = place.get("name", "")
            address = place.get("formatted_address", "")
            types = ", ".join(place.get("types", []))

            # Извлечение координат
            location = place.get("geometry", {}).get("location", {})
            latitude = location.get("lat", "")
            longitude = location.get("lng", "")

            return {
                "status": status,
                "name": name,
                "address": address,
                "types": types,
                "latitude": latitude,
                "longitude": longitude,
            }
        else:
            return {
                "status": "unmatch",
                "name": "",
                "address": "",
                "types": "",
                "latitude": "",
                "longitude": "",
            }
    except Exception as e:
        print(f"Ошибка: {e}")
        return {
            "status": "other",
            "name": "",
            "address": "",
            "types": "",
            "latitude": "",
            "longitude": "",
        }


# Чтение Excel файла
df = pd.read_excel("AcBusTest.xlsx", header=1, engine="openpyxl")
nombres = df["NOMBRE"].dropna().tolist()

# Поиск для всех названий
results = []
for nombre in nombres:
    place_info = check_place(nombre)
    place_info["nombre"] = nombre  # Добавляем оригинальное название организации
    results.append(place_info)
    time.sleep(0.1)  # Небольшая задержка между запросами

# Запись результатов в новый Excel файл
output_df = pd.DataFrame(results)
output_df.to_excel("resultados_organizaciones.xlsx", index=False)

print("Задача выполнена. Результаты сохранены в 'resultados_organizaciones.xlsx'.")
