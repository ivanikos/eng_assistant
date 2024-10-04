import pandas as pd


"""Read the tag-list .xlsm/xlsx etc."""
def read_tags_pd(path_tag_list_pd):
    tag_list = {}

    df_cn = pd.read_excel(path_tag_list_pd, usecols=[2, 4], sheet_name="Support Tags", header=0)
    contract_number = "N/A"
    for row in df_cn.values:
        if (row[0] != "nan" or "n") and row[0] and r"N/A" in contract_number:
            if "Contract Number:" in row[0]:
                contract_number = row[1]
                print(contract_number)

    df = pd.read_excel(path_tag_list_pd, usecols=[2, 5, 11], sheet_name="Support Tags", header=7)
    for row in df.values:
        if row[0] != "nan" or "n":
            load_tag = row[0]
            # contract_number = str(row[1])[2: 10]
            sd_tag = str(row[2]).replace(f"-{contract_number}", "").replace("nan", "")

            tag_list[load_tag] = sd_tag

        else:
            break
    return tag_list
