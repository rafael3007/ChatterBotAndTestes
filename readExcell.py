from datetime import date, timedelta,datetime

# Pega o dia atual
today_date = date.today()
# -35 dias 
td = timedelta(-35)

data_resultante = today_date + td

data_em_texto = '{}/{}/{}'.format(data_resultante.day, data_resultante.month,
data_resultante.year)

print(data_em_texto)