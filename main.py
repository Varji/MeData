import pandas as pd

def extract_data(df, doctor, keywords):

  filtered_df = df[df['name_sotr'] == doctor]
  extracted_data = []

  for index, row in filtered_df.iterrows():
    protocol_text = row['Протокол']

    patient_data = {
      'tkey': row['tkey'],
      'birthday': row['birthday'],
      'd_prm': row['d_prm'],
      'idvisit': row['idvisit'],
      'id_uslug': row['id_uslug'],
      'name_sotr': row['name_sotr'],
    }

    for keyword in keywords:
      if keyword in protocol_text:
        if keyword == 'ПСА':
          psa_value = protocol_text.split(keyword)[1].split('нг')[0].strip()
          patient_data['ПСА'] = psa_value
        elif keyword in ['Гипоэхогенные узлы', 'Инвазия', 'Метастазы (mts)']:
          patient_data[keyword] = protocol_text.split(keyword)[1].strip()
        else:
          print(f'Неизвестное ключевое слово: {keyword}')

    extracted_data.append(patient_data)

  return pd.DataFrame(extracted_data)

df = pd.read_excel('c61_fio1_uslugs.xls')
doctor = 'ВОЛКОВ ВЛАДИСЛАВ МИХАЙЛОВИЧ'
keywords = ['Гипоэхогенные узлы', 'Инвазия', 'Метастазы (mts)', 'ПСА']
extracted_data = extract_data(df, doctor, keywords)
print(extracted_data.to_string(index=False))