
#powered Alexander Andrade


import gspread
from google.auth import exceptions
from google.auth.transport.requests import Request
from google.oauth2 import service_account

# Configuration API
credentials_path = "projeto-devops-413718-80f22797628f.json"
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# json credentials
try:
    credentials = service_account.Credentials.from_service_account_file(credentials_path, scopes=scope)
except exceptions.GoogleAuthError:
    # aut attempts
    credentials = credentials.refresh(Request())

gc = gspread.Client(auth=credentials)

# spreadsheet url
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1eCxt_auWF2n5uvkd2IGccwaXREZeQxC5AZfxQTrukVE/edit#gid=0'
sh = gc.open_by_url(spreadsheet_url)

# head
worksheet = sh.get_worksheet(0)

data = worksheet.get_all_values()

# call to lines
for i, row in enumerate(data[3:], start=4):
   
    if len(row) >= 6:
        # columns
        matricula, nome, faltas, p1, p2, p3 = map(str, row[:6])

        # avg
        total_aulas = 60

        # sum avg
        media = (float(p1) + float(p2) + float(p3)) / 3


        # student situation

        if int(faltas) > 0.25 * total_aulas:
            situacao = 'Reprovado por Falta'
            worksheet.update_cell(i, 8, 0)
        elif media < 5:
            situacao = 'Reprovado por Nota'
            worksheet.update_cell(i, 8, 0)
        elif 5 <= media < 7:
            situacao = "Exame Final"
            naf = max(0, 2 * (7 - media))
            naf = int(naf + 0.5)
            worksheet.update_cell(i, 8, naf)
            # calculate score for (naf)
            naf = 5 <= (media)/2
            # round whole number  
            naf = int(naf + 0.5)
            # udpate cell
            worksheet.update_cell(i, 7, naf)
        else:
            situacao = 'Aprovado'
            # update cell
            worksheet.update_cell(i, 8, 0)

        # update situation
        worksheet.update_cell(i, 7, situacao)


# powered Alexander Silva de Andrade