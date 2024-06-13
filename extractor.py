import os.path

from pypdf import PdfReader
from openpyxl import Workbook


def extract(file):
    reader = PdfReader(file)
    extracted_list = []
    for page in reader.pages:
        unformatted_text = page.extract_text(extraction_mode="layout")
        extracted_list.extend(unformatted_text.splitlines())
    return extracted_list


def readfile():
    supp = os.path.join("C:/Users/tweyt/PycharmProjects/SupplementalBid", "cdb3be27-6799-4ce2-aea3-284edd536c33.pdf")
    sen_list = os.path.join("C:/Users/tweyt/PycharmProjects/SupplementalBid", "alpa_q2_2024.pdf")
    return supp, sen_list


def supplemental_to_dict(raw_list):
    del raw_list[0]
    formatted_list = []
    for line in raw_list:
        text = line.split()
        last_name = text[2]
        data = {
            "EmpID": text[0],
            "Seniority": text[1],
            "LastName": last_name.rstrip(','),
            "EffDate": text[-1],
            "BidType": text[-2],
            "Awarded": text[-3]
        }

        given_names = []
        for i in range(-4, -10, -1):
            if text[i] == last_name:
                break
            if text[i].endswith(','):
                data["LastName"] += text[i]
            else:
                given_names.append(" " + text[i])
        given_names.reverse()
        data["GivenNames"] = given_names
        formatted_list.append(data)
    return formatted_list


def seniority_to_dict(raw_sen):
    del raw_sen[0]
    formatted_list = []
    for line in raw_sen:
        text = line.split()
        data = {
            "EmpID": text[0],
            "YRS2RTR": text[-1],
            "RTRDate": text[-2],
            "HIREDATE": text[-3],
            "SEN": text[-4],
            "SEAT": text[-5],
            "FLEET": text[-6],
            "BASE": text[-7]
        }
        given_names = []
        for i in range(-8, -20, -1):
            if len(text) + i < 0:
                break
            if text[i] == text[0]:
                break
            else:
                given_names.append(" " + text[i])
        given_names.reverse()
        data["FullNames"] = given_names
        formatted_list.append(data)
    return formatted_list


def load_and_convert_to_dicts():
    supp, sen_list = readfile()
    raw_supp = extract(supp)
    raw_sen = extract(sen_list)
    fl = supplemental_to_dict(raw_supp)
    sl = seniority_to_dict(raw_sen)
    return fl, sl


def generate_indexed_dicts(fl, sl):
    supp = {x["EmpID"]: x for x in fl}
    sen = {x["EmpID"]: x for x in sl}
    return supp, sen


def combiner():
    fl, sl = load_and_convert_to_dicts()
    supp, sen = generate_indexed_dicts(fl, sl)
    combined = {}
    for x in sen.keys():
        combined[x] = sen[x]
        y = supp.get(x)
        if y:
            combined[x].update(y)
    leftovers = supp.keys() - sen.keys()
    for x in leftovers:
        combined.update(supp[x])

    # combined = {x: sen.get(x).update(supp.get(x)) for x in sen.keys() if supp.get(x)}
    return combined


def create_spreadsheet(data):
    wb = Workbook()
    ws = wb.active
    ws.append(data.keys())
    for row in data.values():
        ws.append(row)
    return ws
