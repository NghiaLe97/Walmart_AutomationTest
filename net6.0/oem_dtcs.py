import json
import os
import re
import pandas as pd


log_folder = r"C:\Users\nghia\PycharmProjects\Walmart_AutomationTest\Database\Log_file"
path_file = r"C:\Users\nghia\PycharmProjects\Walmart_AutomationTest\Database\Test Document\All Make_Test Document_V0.16_May222024.xlsx"
def longest_list(*lists):
    # Tìm danh sách dài nhất trong các danh sách được truyền vào
    longest = max(lists, key=len)
    return longest

def oem_dtcs_expected(sys):
    df = pd.read_excel(path_file, sheet_name="Jaguar Land Rover")
    data_system = df.loc[df["VIN"] == "SAJAK4BVXHCP15541"]
    data_system_2 = data_system.loc[data_system["Functions/SubFunctions"] == "NWS"]
    data_system_3 = data_system_2.loc[df["Value (US)"] != "Not Support"]
    data_system_4 = data_system_3.loc[df['7111 7" Android VCI'] == "v"]
    data_system_5 = data_system_4.loc[data_system_4["System/SubSystem"] == "{0}".format(sys)]
    sliced_data = data_system_5.copy()
    sliced_data['combine'] = sliced_data['DTC'].astype(str) + '-' + sliced_data['Status'].astype(str)
    sliced_data['combine'] = sliced_data['combine'].apply(lambda x: x.lower().replace('|', '/').replace(' ', ''))
    sliced_data['Value (US)'] = sliced_data['Value (US)'].str.lower()
    dtc_def_document = sliced_data.set_index('combine')['Value (US)'].to_dict()

    filtered_dict = {}
    for key, value in dtc_def_document.items():
        if key.endswith('-nan'):
            new_key = key.rsplit('-nan', 1)[0]
            filtered_dict[new_key] = value.strip()
        else:
            filtered_dict[key] = value.strip()

    return filtered_dict

a = oem_dtcs_expected("ATCM-All Terrain Control Module")
print(a)


def systems_list_excel():
    systems_list = []
    df = pd.read_excel(path_file, sheet_name="Jaguar Land Rover")
    data_system = df.loc[df["VIN"] == "SAJAK4BVXHCP15541"]
    data_system_2 = data_system.loc[data_system["Functions/SubFunctions"] == "NWS"]
    data_system_3 = data_system_2.loc[df["Value (US)"] != "Not Support"]
    data_system_4 = data_system_3.loc[df['7111 7" Android VCI'] == "v"]
    data_system_4_tolist = data_system_4["System/SubSystem"].tolist()
    for i in data_system_4_tolist:
        i_lower = i.lower().strip()
        if i_lower not in systems_list:
            systems_list.append(i_lower)
    return systems_list

b = systems_list_excel()
print(b)

for filename in os.listdir(log_folder):
    if filename.endswith(".txt"):  # Đảm bảo là tệp tin văn bản
        file_path = os.path.join(log_folder, filename)

        # Đọc nội dung tệp tin
        with open(file_path, 'r') as file:
            log_data = file.read()

        def get_oem_dtcs():

            # Regex để tìm [oemModuleDtcs]
            pattern = r"\[oemModuleDtcs\]: (.*?)\n"

            # Tìm các phần khớp
            matches = re.findall(pattern, log_data, re.DOTALL)

            if matches:  # Kiểm tra nếu có kết quả
                # Tìm phần tử dài nhất trong matches
                oem_dtcs = longest_list(*matches)
                # print(oem_dtcs)
                # print(json.dumps(json.loads(oem_dtcs), indent=4))
                return oem_dtcs
            else:
                print(f"No matches found in {filename}")

def compare_lists(system, sub_system, document, actual):
    # Find elements in document that are not in actual
    only_in_document = [item for item in document if item not in actual]

    # Find elements in actual that are not in document
    only_in_actual = [item for item in actual if item not in document]

    # Find elements that are common in both lists
    common_elements = [item for item in document if item in actual]


    # Open the CSV file and write the results
    with open(csv_file_path, 'a', newline='') as f:
        writer = csv.writer(f)

        # Write total number of elements in each list
        new_row = [system, sub_system, 'total DTC/PIDs in document', len(document), 'total DTC/PIDs in actual', len(actual), 'NA']
        writer.writerow(new_row)

        # Write elements only found in document
        for item in only_in_document:
            writer.writerow([system, sub_system, item, 'null', 'null', 'null', 'Fail - only in document'])

        # Write elements only found in actual
        for item in only_in_actual:
            writer.writerow([system, sub_system, 'null', 'null', item, 'null', 'Fail - only in app'])

        # Write common elements
        for item in common_elements:
            writer.writerow([system, sub_system, item, 'null', item, 'null', 'Pass'])