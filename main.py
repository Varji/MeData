import pandas as pd

df = pd.read_excel("c61_fio1_uslugs.xls")


def find_data_by_doctor(doctor_names):
    doctor_names = doctor_names.split(",")
    doctor_df = df[df["name_sotr"].str.contains('|'.join(doctor_names), case=False, na=False)]
    doctor_df = doctor_df.sort_values(by=['tkey', 'd_prm'])

    data = []
    for tkey in doctor_df['tkey'].unique():
        row = doctor_df[doctor_df['tkey'] == tkey].iloc[-1]

        psa = None
        date = None
        volume = None
        hypoechoic_nodes = "не обнаружены"
        invasion = "не выявлено"
        invasion_location = None
        mts = "не выявлено"
        mts_location = None

        protocol = row["Протокол"].lower() 

        if "пса" in protocol:
            psa_start = protocol.find("пса")
            psa_end = protocol.find("нг", psa_start)
            psa = row["Протокол"][psa_start:psa_end].strip()

        date = row["d_prm"]

        if "объем" in protocol:
            volume_start = protocol.find("объем")
            volume_end = protocol.find("см3", volume_start)
            volume = row["Протокол"][volume_start:volume_end].strip()

        if "гипоэхоген" in protocol:
            hypoechoic_nodes = "обнаружены"

        if "инвази" in protocol:
            if "не выявлено" in protocol[protocol.find("инвази"):]:
                invasion = "не выявлено"
            else:
                invasion = "выявлено"
                invasion_start = protocol.find("инвази") + 7
                invasion_end = protocol.find(".", invasion_start)
                invasion_location = row["Протокол"][invasion_start:invasion_end].strip() if invasion_end != -1 else row["Протокол"][
                                                                                                invasion_start:].strip()

        if "mts" in protocol:
            if "не выявлено" in protocol[protocol.find("mts"):]:
                mts = "не выявлено"
            else:
                mts = "выявлено"
                mts_start = protocol.find("mts") + 3
                mts_end = protocol.find(".", mts_start)
                mts_location = row["Протокол"][mts_start:mts_end].strip() if mts_end != -1 else row["Протокол"][
                                                                                                mts_start:].strip()
        if any([psa, volume, hypoechoic_nodes == "обнаружены", invasion == "выявлено", mts == "выявлено"]):
            data.append({
                "tkey": row['tkey'],
                "idvisit": row['idvisit'],
                "id_uslug": row['id_uslug'],
                "ПСА": psa,
                "Дата": date,
                "Объем": volume,
                "Гипоэхогенные л/узлы": hypoechoic_nodes,
                "Инвазия за капсулу": invasion,
                "Локализация инвазии": invasion_location,
                "MTS": mts,
                "Локализация MTS": mts_location
            })

    result_df = pd.DataFrame(data)

    return result_df


result_df = find_data_by_doctor("Волков В.М.,ВОЛКОВ ВЛАДИСЛАВ МИХАЙЛОВИЧ")

result_df.to_excel("volkov_data.xlsx", index=False)