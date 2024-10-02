import requests
from bs4 import BeautifulSoup
import re
from collections import defaultdict
from datetime import datetime
import openpyxl

# URL
r = requests.get('')
soup = BeautifulSoup(r.text, 'html.parser')

start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 8, 31)

player_names = [
    'Alisa', 'Asuka', 'Azucena', 'Bryan', 'Claudio', 'Devil Jin', 'Dragunov', 'Feng',
    'Hwoarang', 'Jack-8', 'Jin', 'Jun', 'Kazuya', 'King', 'Kuma', 'Lars', 'Lee',
    'Leo', 'Leroy', 'Lili', 'Xiaoyu', 'Law', 'Nina', 'Panda', 'Paul', 'Raven',
    'Reina', 'Shaheen', 'Steve', 'Victor', 'Yoshimitsu', 'Zafina', 'Eddy', 'Lidia'
]

categorized_data = defaultdict(list)
statistics = defaultdict(lambda: defaultdict(lambda: {'WIN': 0, 'LOSE': 0, 'DRAW': 0, 'total': 0}))
total_statistics = defaultdict(lambda: {'WIN': 0, 'LOSE': 0, 'DRAW': 0, 'total': 0})

def clean_text(text):
    text = text.replace('\n', ' ')
    return re.sub(r'\s+', ' ', text).strip()

def convert_datetime_format(datetime_str):
    try:
        datetime_obj = datetime.strptime(datetime_str, '%d %b %Y %H:%M')
        return datetime_obj.strftime('%Y-%m-%d %H:%M')
    except ValueError:
        return datetime_str

def is_within_date_range(date_str):
    try:
        date_obj = datetime.strptime(date_str, '%d %b %Y %H:%M')
        return start_date <= date_obj <= end_date
    except ValueError:
        return False

if len(soup.find_all('tbody')) >= 2:
    tbody = soup.find_all('tbody')[1]
    rows = tbody.find_all('tr')

    for row in rows:
        cells = row.find_all('td')
        cell_values = [clean_text(cell.get_text()) for cell in cells]

        if len(cell_values) > 1:
            datetime_str = cell_values[0]
            if is_within_date_range(datetime_str):
                converted_datetime_str = convert_datetime_format(datetime_str)
                cell_values[0] = converted_datetime_str

                key = cell_values[1]
                categorized_data[key].append(cell_values)

                opponent = cell_values[5]
                result = cell_values[2].split()[1]
                if opponent in player_names:
                    statistics[key][opponent][result] += 1
                    statistics[key][opponent]['total'] += 1

                    total_statistics[opponent][result] += 1
                    total_statistics[opponent]['total'] += 1

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'player_statistics_{timestamp}.xlsx'

    workbook = openpyxl.Workbook()
    total_sheet = workbook.create_sheet(title='Total')

    for player in player_names:
        if player in statistics:
            sheet = workbook.create_sheet(title=player)
            headers = ['Opponent', 'Wins', 'Losses', 'Draws', 'Win Rate']
            sheet.append(headers)

            total_wins = total_losses = total_draws = total_matches = 0

            for opponent in player_names:
                if opponent in statistics[player]:
                    stats = statistics[player][opponent]
                    wins = stats['WIN']
                    losses = stats['LOSE']
                    draws = stats['DRAW']
                    total = stats['total']
                    win_rate = (wins / total) * 100 if total > 0 else 0
                    row = [opponent, wins, losses, draws, f'{win_rate:.2f}%']
                    sheet.append(row)
                    total_wins += wins
                    total_losses += losses
                    total_draws += draws
                    total_matches += total

            sheet.append(['Total', total_wins, total_losses, total_draws, f'{(total_wins / total_matches) * 100 if total_matches > 0 else 0:.2f}%'])

    total_headers = ['Opponent', 'Wins', 'Losses', 'Draws', 'Win Rate']
    total_sheet.append(total_headers)

    grand_total_wins = grand_total_losses = grand_total_draws = grand_total_matches = 0

    for opponent in player_names:
        if opponent in total_statistics:
            stats = total_statistics[opponent]
            wins = stats['WIN']
            losses = stats['LOSE']
            draws = stats['DRAW']
            total = stats['total']
            win_rate = (wins / total) * 100 if total > 0 else 0
            total_sheet.append([opponent, wins, losses, draws, f'{win_rate:.2f}%'])
            grand_total_wins += wins
            grand_total_losses += losses
            grand_total_draws += draws
            grand_total_matches += total

    total_sheet.append(['Total', grand_total_wins, grand_total_losses, grand_total_draws, f'{(grand_total_wins / grand_total_matches) * 100 if grand_total_matches > 0 else 0:.2f}%'])

    if 'Sheet' in workbook.sheetnames:
        workbook.remove(workbook['Sheet'])

    workbook.save(filename)

    for player, matches in categorized_data.items():
        txt_filename = f'{player}.txt'
        with open(txt_filename, 'w', encoding='utf-8') as txtfile:
            for match in matches:
                txtfile.write('\t'.join(match) + '\n')

else:
    print('End of Data.')
