from pandas import read_csv, read_excel, merge
from datetime import datetime

def getSpreadSheets(_input):
    ssInfo = read_csv(_input, sep='\t',encoding='utf-16', on_bad_lines='skip', nrows=6)
    ssIn = read_csv(_input, sep='\t',encoding='utf-16', on_bad_lines='skip',skiprows=7)
    return ssIn, ssInfo

def convertToDatetime(lst):
    date = []
    for line in lst:
        date.append(datetime.strptime(line, '%m/%d/%Y, %I:%M:%S %p'))
    return date

def convertDuration(duration):
    return list(map(lambda x: x.total_seconds(),duration))

def getClassDate(df):
    dt = convertToDatetime([df.loc['Hora de início da reunião']['Resumo da Reunião']])
    return dt[0].strftime('%d/%m/%Y')

def isPresent(spreadsheet):
    presenca = []
    percent = 0.4
    minDuration = spreadsheet['Duração'].mean() * percent
    for d in spreadsheet['Duração']:
        if d >= minDuration:
            presenca.append('1')
        else:
            presenca.append('0')
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
        ssOut[date].fillna('0', inplace=True)
        ssOut.to_excel(output,index=False)
        return True
    except:
        return False

def dataProcessing(ss):
    # Descartar colunas
    ss.drop(columns=['Email', 'Função'], inplace=True)
    # Ordenação por nome
    ss.sort_values(by='Nome Completo', inplace=True)
    # Obtendo a matricula
    ss.rename(columns={'ID do participante (UPN)': 'Matrícula'}, inplace=True)
    ss['Matrícula'] = getId(ss['Matrícula'])
    # Conversão de datas para o formato datetime
    ss['Horário de Entrada']= convertToDatetime(ss['Horário de Entrada'])
    ss['Horário de Saída'] = convertToDatetime(ss['Horário de Saída'])
    ss['Duração'] = convertDuration(ss['Horário de Saída'] - ss['Horário de Entrada'])
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

    return generateAttendenceSheet(ss[['Matrícula', classDate]], path_out, classDate)
