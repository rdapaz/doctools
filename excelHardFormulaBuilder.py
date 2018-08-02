
rows = [23,44,109,59,123]

start_row = 11

for idx, row in enumerate(rows):
    formulas = [
                f'=Budget!E{row}',
                f'=C{start_row+idx}*0.15',
                f'=SUM(C{start_row+idx}+D{start_row+idx})',
                '',
                f'=SUM(INDIRECT("\'Forecast\'!G{row}:"&$F$5&ROW(Forecast!G{row})))',
                f'=SUM(Forecast!G{row}:R{row})+SUM(Forecast!T{row}:AE{row})-G{start_row+idx}'
                ]
    print('|'.join(formulas))

