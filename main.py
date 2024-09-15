import pandas as pd

df = pd.read_excel("c61_fio1_uslugs.xls")

def find_data_by_doctor(doctor_names):
    doctor_names = doctor_names.split(",")
    doctor_df = df[df["name_sotr"].str.contains('|'.join(doctor_names), case=False, na=False)]
    data = []
    for index, row in doctor_df.iterrows():
        psa = None
        date = None
        volume = None
        hypoechoic_nodes = None
        invasion = None
        mts = None

        if "ПСА" in row["Протокол"]:
            psa_start = row["Протокол"].find("ПСА")
            psa_end = row["Протокол"].find("нг", psa_start)
            psa = row["Протокол"][psa_start:psa_end].strip()

        date = row["d_prm"]

        if "Объем" in row["Протокол"]:
            volume_start = row["Протокол"].find("Объем")
            volume_end = row["Протокол"].find("см3", volume_start)
            volume = row["Протокол"][volume_start:volume_end].strip()

        if "гипоэхоген" in row["Протокол"].lower():
            hypoechoic_nodes = True

        if "инвази" in row["Протокол"].lower():
            invasion = True

        if "mts" in row["Протокол"].lower():
            mts = True

        data.append({
            "ПСА": psa,
            "Дата": date,
            "Объем": volume,
            "Гипоэхогенные л/узлы": hypoechoic_nodes,
            "Инвазия за капсулу": invasion,
            "MTS": mts
        })

    result_df = pd.DataFrame(data)

    return result_df

result_df = find_data_by_doctor("Волков В.М.,ВОЛКОВ ВЛАДИСЛАВ МИХАЙЛОВИЧ")

result_df.to_excel("volkov_data.xlsx", index=False)