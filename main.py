from bs4 import BeautifulSoup, Tag
from openpyxl import Workbook
from datetime import datetime
import os, glob, sys

current_directory = os.path.dirname(os.path.abspath(__file__))
html_files = glob.glob(os.path.join(current_directory, '*.html'))

leaderboard_data = []

for html_file in html_files:
    f = open(html_file, "r")

    soup = BeautifulSoup(f.read(), 'html.parser')

    table = soup.find('table', id='leaderBoard__TablePlayers')
    if not isinstance(table, Tag): break

    tbody = table.find("tbody")
    if not isinstance(tbody, Tag): break

    rows = tbody.find_all("tr")
    for tr in rows:
        pos = tr.find('td', class_='leaderBoard__Table__Position').text.strip()
        name = tr.find('td', class_='leaderBoard__Table__Name').find("span").text.strip()
        points = tr.find('div', class_='leaderBoard__Table__Score__Points').find("div").text.strip()

        leaderboard_data.append({
            'pos': pos,
            'name': name,
            'points': points
        })

    f.close()

sorted_data = sorted(leaderboard_data, key=lambda x: x['name'].lower())

wb = Workbook()
ws = wb.active

if ws is None:
    sys.exit()

ws.title = "Leaderboard"
ws.append(["Позиция", "Имя", "Очки"])
for entry in sorted_data:
    ws.append([entry['pos'], entry['name'], entry['points']])

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
excel_file_path =f"leaderboard_{timestamp}.xlsx" 
wb.save(excel_file_path)

print(f"Файл {excel_file_path} успешно сохранен.")
