from openpyxl import load_workbook
import pandas as pd
import os
import json

path = r"C:\Users\enzoc\Desktop\dados_inversores_espaçoPlaneta\inversor2"
inversor = os.path.basename(path)
list_of_files = os.listdir(path)

count = 0
list_final = []

for i in list_of_files:
    count += 1
    path_final = path + r"\\" + i
    path_without_extension = os.path.basename(path_final)[:-19]
    path_save = r"./files/" + path_without_extension + r'.xlsx'
    if inversor == "inversor1":
        inverter_sn = path_without_extension[10:]
        inverter_sn = inverter_sn[:15]
    if inversor == "inversor2":
        inverter_sn = path_without_extension[:15]

    excel = load_workbook(path_final)
    sheet = excel['Sheet0']
    sheet.delete_rows(1,3)

    if inversor == "inversor1":
        sheet.delete_cols(2,8)
        sheet.delete_cols(3,6)
        sheet.delete_cols(11,1)
        sheet.delete_cols(9,1)
        sheet.delete_cols(4,1)
        sheet.delete_cols(10,1)

    if inversor == "inversor2":
        sheet.delete_cols(2,12)
        sheet.delete_cols(3,6)
        sheet.delete_cols(11,1)
        sheet.delete_cols(9,1)
        sheet.delete_cols(4,1)
        sheet.delete_cols(10,1)

    excel.save(path_save)
    df = pd.read_excel(path_save)

    df.rename(
        columns = {
                    "Time":"time",
                    "DC Voltage PV1(V)":"dc_voltage_pv1",
                    "DC Voltage PV2(V)":"dc_voltage_pv2",
                    "DC Current1(A)":"dc_current1",
                    "DC Current2(A)":"dc_current2",
                    "Total DC Input Power(W)":"total_dc_input_power",
                    r"AC Voltage R/U/A(V)":"ac_voltage_r_u_a",
                    r"AC Voltage S/V/B(V)":"ac_voltage_s_v_b",
                    r"AC Voltage T/W/C(V)":"ac_voltage_t_w_c",
                    r"AC Current R/U/A(A)":"ac_current_r_u_a",
                    r"AC Current S/V/B(A)":"ac_current_s_v_b",
                    r"AC Current T/W/C(A)":"ac_current_t_w_c",
                    "AC Output Total Power (Active)(W)":"ac_output_total_power",
                    "AC Output Frequency R(Hz)":"ac_output_frequency_R",
                    "Generation of Last Month (Active)(kWh)":"generation_of_last_month",
                    "Daily Generation (Active)(kWh)":"daily_generation",
                    "Total Generation (Active)(kWh)":"total_generation",
                    "Power Grid Total Apparent Power(VA)":"power_grid_total_apparent_power",
                    "Grid Power Factor":"grid_power_factor",
                    "Inverter Temperature(℃)":"inverter_temperature",
                    "Inverter Status":"inverter_status",
                    "Generation Yesterday(kWh)":"generation_yesterday",
                    "System Time":"system_time"
                },
        inplace=True
    )
    df['time'] = df['time'].astype(str)
    df = df.to_dict("records")
    list_final.extend(df)

data = {
  "inverter_sn": inverter_sn,
  "data": list_final,
}

with open(f'{inverter_sn}.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
