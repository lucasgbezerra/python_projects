from pandas import read_csv, read_excel, merge
from datetime import datetime

def getSpreadSheets(_input):
    global typeCsv
    ssInfo = read_csv(_input, sep='\t',encoding='utf-16', on_bad_lines='skip', nrows=7)
    typeCsv = getTypeCsv(ssInfo)
    if typeCsv == 0:
        ssIn = read_csv(_input, sep='\t',encoding='utf-16', on_bad_lines='skip',skiprows=7)
    else:
        ssIn = read_csv(_input, sep='\t',encoding='utf-16', on_bad_lines='skip',skiprows=8, nrows=ssInfo.iloc[0][0])
    return ssIn, ssInfo

def convertToDatetime(lst):
    date = []
    for line in lst:
        if type(line) is str:
            if typeCsv == 1:
                line = datetime.strptime(line, '%m/%d/%y, %I:%M:%S %p')
            else:
                line = datetime.strptime(line, '%m/%d/%Y, %I:%M:%S %p')
        date.append(line)
    return date

def getTypeCsv(df):
    global horaEntrada
    global horaSaida
    global duracao
    if df.columns[0] == '1. Resumo':
        # Tipo com 2 tabelas
        horaEntrada = 'Primeiro ingresso'
        horaSaida = 'Última saída'
        duracao = 'Duração da\xa0reunião'
        return 1
    else:
        # Tipo com 1 tabela
        
        horaEntrada = 'Horário de Entrada'
        horaSaida = 'Horário de Saída'
        duracao = 'Duração'
        return 0

def convertDuration(duration):
    return list(map(lambda x: x.total_seconds(),duration))

    
def getClassDate(df):
    dt = df.iloc[2][0]
    if typeCsv == 1:
        converted = datetime.strptime(dt, '%m/%d/%y, %I:%M:%S %p')
    else:
        converted = datetime.strptime(dt, '%m/%d/%Y, %I:%M:%S %p')
    return f"{converted.strftime('%d')}/{converted.strftime('%m')}"

def isPresent(spreadsheet):
    presenca = []
    percent = 0.4
    minDuration = spreadsheet['Duração'].mean() * percent
    for d in spreadsheet['Duração']:
        if d >= minDuration:
            presenca.append(int(1))
        else:
            presenca.append(int(0))
    return presenca

def getId(lst):
     return list(map(lambda x: x.split('@')[0],lst))

def totalTime(df):
    result = merge(df.groupby('Matrícula').sum(), df.loc[:,df.columns != 'Duração'], how='left', on='Matrícula')
    result.drop_duplicates(subset=['Matrícula'],inplace=True)
    return result

def generateAttendenceSheet(ssIn, output, date):
    try:
        ssOut = read_excel(output)
        ssOut = ssOut.astype({'Matrícula': str})
        ssOut = merge(ssOut, ssIn, how='left',on='Matrícula')
        ssOut[date].fillna(int(0), inplace=True)
        ssOut = ssOut.astype({date: int})
        ssOut.to_excel(output,index=False)
        return True
    except:
        return False
    
def dataProcessing(ss):
    ss.drop(columns=['Email', 'Função',duracao], inplace=True)        # Ordenação por nome
    if typeCsv == 1:
        ss.sort_values(by="Nome", inplace=True)
    else:
        ss.sort_values(by="Nome Completo", inplace=True)
    # Obtendo a matricula
    ss.rename(columns={'ID do participante (UPN)': 'Matrícula'}, inplace=True)
    ss['Matrícula'] = getId(ss['Matrícula'])
    # Conversão de datas para o formato datetime
    ss[horaEntrada]= convertToDatetime(ss[horaEntrada])
    ss[horaSaida] = convertToDatetime(ss[horaSaida])
    ss['Duração'] = convertDuration(ss[horaSaida] - ss[horaEntrada])
    ss = totalTime(ss)
    # Remoção de Não alunos
    ss.set_index('Matrícula', inplace=True)
    ss.drop('elainevenson', inplace=True)
    return ss

def main(path_in, path_out):

    ss, ssInfo = getSpreadSheets(path_in)
    classDate = getClassDate(ssInfo)

    ss = dataProcessing(ss)

    ss[classDate] = isPresent(ss)
    ss.reset_index(inplace=True)
#     return ss, classDate
    return generateAttendenceSheet(ss[['Matrícula', classDate]], path_out, classDate)
