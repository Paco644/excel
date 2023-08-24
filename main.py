import os.path

import gradio as gr
import shutil

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook, Workbook


def increment_id(last_id):
    return last_id[:-4] + '0' * (4 - len(t := str(int(last_id.split('-')[2]) + 1))) + t


def transfer_data(price_list, shop_list, progress=gr.Progress()):
    if price_list is None:
        raise Exception("Bitte beide Datein hochladen!")

    if shop_list is None:
        raise Exception("Bitte beide Datein hochladen!")

    shutil.copy(price_list.name, "price_list.xlsx")
    shutil.copy(shop_list.name, "shop_list.xlsx")

    price_wb: Workbook = load_workbook("price_list.xlsx")
    shop_wb: Workbook = load_workbook("shop_list.xlsx")

    # Get all already known ids
    try:
        shop_ws: Worksheet = shop_wb['Artikel']
    except KeyError:
        raise Exception("Shopliste enthält falsche daten, vielleicht vertauscht?")

    known_articles = {}

    shop_data = shop_ws['B:D']
    ids = shop_data[0]
    name = shop_data[2]

    for i in range(3, len(ids)):
        _id = ids[i].value
        _name = name[i].value
        if _id is None or _name is None:
            continue
        _name = _name.strip()
        known_articles[_id] = _name
    try:
        price_ws: Worksheet = price_wb['40495396']
    except KeyError:
        raise Exception("Preisliste enthält falsche daten, vielleicht vertauscht?")

    new_articles = {}

    price_data = price_ws['A:C']
    name = price_data[0]
    price = price_data[2]

    for i in range(1, len(name) - 5):
        _name = name[i].value
        _price = price[i].value
        if _name == "Summe" or _name is None or _price is None:
            continue
        _name = _name.strip()
        # Check if name exists in known articles
        if _name not in known_articles.values():
            new_articles[_name] = _price

    print(new_articles)

    if len(new_articles) > 0:

        last_id = str(list(known_articles)[-1])
        new_id = increment_id(last_id)
        total_rows = len(ids) + 1

        for new_article in progress.tqdm(new_articles, desc="Inserting rows...", unit="rows"):
            shop_ws.cell(row=total_rows, column=2).value = new_id
            shop_ws.cell(row=total_rows, column=4).value = new_article
            shop_ws.cell(row=total_rows, column=5).value = "BITTE AUSFÜLLEN"
            shop_ws.cell(row=total_rows, column=8).value = "Circ IT"
            shop_ws.cell(row=total_rows, column=12).value = "Stück"
            shop_ws.cell(row=total_rows, column=16).value = new_articles[new_article]
            shop_ws.cell(row=total_rows, column=20).value = "EUR"
            shop_ws.cell(row=total_rows, column=21).value = "19"

            shop_ws.cell(row=total_rows, column=16).number_format = '#,##0.00€'
            shop_ws.cell(row=total_rows, column=21).number_format = "00"
            # update values
            new_id = increment_id(new_id)
            total_rows += 1

        shop_wb.save("new_shop_list.xlsx")
        return gr.update(value=os.path.abspath("new_shop_list.xlsx"))
    else:
        raise Exception("Nothing new to add...")


with gr.Blocks(theme=gr.themes.Soft()) as app:
    with gr.Row():
        price_list = gr.File(label="Preisliste")
        shop_list = gr.File(label="Shopliste")

    send_button = gr.Button("Übertragen")

    output = gr.File(label="Neue Shopliste")

    send_button.click(transfer_data, inputs=[price_list, shop_list], outputs=[output])

app.launch(show_error=True, server_port=8080, ssl_verify=True, enable_queue=True)
