from os import path, system

import gradio as gr
import shutil

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook, Workbook

system("git pull")


class Product:
    def __init__(self, name: str, price: float):
        self.id = -1
        self.name = name
        self.price = price


class Bundle:
    def __init__(self, name: str, products=None):
        if products is None:
            products: list[Product] = []
        self.name = name
        self.products = products
        self.sum = -1

    def add_product(self, product: Product):
        self.products.append(product)

    def set_sum(self, sum: float):
        self.sum = sum

    def calculate_sum(self):
        for p in self.products:
            self.sum += p.price
        self.sum += 1


def increment_id(last_id):
    return last_id[:-4] + '0' * (4 - len(t := str(int(last_id.split('-')[2]) + 1))) + t


def transfer_data(price_list, shop_list, progress=gr.Progress()):
    if price_list is None:
        raise Exception("Bitte beide Datein hochladen.")

    if shop_list is None:
        raise Exception("Bitte beide Datein hochladen.")

    shutil.copy(price_list.name, "price_list.xlsx")

    price_wb: Workbook = load_workbook("price_list.xlsx")
    shop_wb: Workbook = load_workbook("shop_template.xlsx")

    try:
        price_ws: Worksheet = price_wb['40495396']
        shop_ws: Worksheet = shop_wb['Artikel']
    except KeyError:
        raise Exception("Preisliste enthält falsche daten.")

    price_data = price_ws['A:D']
    name = price_data[0]
    price = price_data[3]
    bundle = None
    new_id = increment_id("HW-CIT-0000")
    total_rows = 4

    bundles: list[Bundle] = []
    products: list[Product] = []

    for i in progress.tqdm(range(1, len(name) - 5), unit="Rows", desc="Getting Pricelist data..."):
        _name = name[i].value
        _price = price[i].value

        if _name is None:
            continue
        if _name == "Summe":
            bundle.calculate_sum()
            bundles.append(bundle)
            bundle = None
            continue
        if _price is None:
            bundle: Bundle = Bundle(_name)
            continue

        product = Product(_name, _price)

        if bundle:
            bundle.add_product(product)
        else:
            products.append(product)

    # Alle Einzelprodukte zu einer Liste zusammenfügen
    for bundle in progress.tqdm(bundles, unit="Bundles", desc="Merging all Products..."):
        products += bundle.products

    # Alle Duplikate entfernen
    scanned = []
    for product in progress.tqdm(products, unit="Products", desc="Removing Duplicates...."):
        if product.name not in scanned:
            scanned.append(product.name)
        else:
            products.remove(product)

    # IDS für die Einzelrprodukte setzen
    for product in progress.tqdm(products, unit="Products", desc="Generating IDs..."):
        product.id = new_id
        new_id = increment_id(new_id)

    # Einzelprodukte in die Excel hinzufügen
    for product in progress.tqdm(products, unit="Products", desc="Inserting Products into excel sheet..."):
        shop_ws.cell(row=total_rows, column=2).value = product.id
        shop_ws.cell(row=total_rows, column=4).value = product.name
        shop_ws.cell(row=total_rows, column=8).value = "Circ IT"
        shop_ws.cell(row=total_rows, column=9).value = "20"
        shop_ws.cell(row=total_rows, column=12).value = "Stück"
        shop_ws.cell(row=total_rows, column=16).value = product.price
        shop_ws.cell(row=total_rows, column=20).value = "EUR"
        shop_ws.cell(row=total_rows, column=21).value = "19"
        shop_ws.cell(row=total_rows, column=21).value = str(product.id) + ".png"
        shop_ws.cell(row=total_rows, column=25).value = "png"

        shop_ws.cell(row=total_rows, column=16).number_format = '#,##0.00€'
        shop_ws.cell(row=total_rows, column=21).number_format = "00"

        total_rows += 1

    # Alle Bundles durchgehen
    for bundle in progress.tqdm(bundles, unit="Bundles", desc="Inserting Bundles into excel sheet..."):

        ids = ""

        # IDs der Einzelrpodukte bekommen
        for bundle_product in bundle.products:
            for product in products:
                if product.name == bundle_product.name:
                    ids += "\n" + str(product.id)

        shop_ws.cell(row=total_rows, column=1).value = "SET"
        shop_ws.cell(row=total_rows, column=2).value = new_id
        shop_ws.cell(row=total_rows, column=4).value = bundle.name
        shop_ws.cell(row=total_rows, column=6).value = "enthält Artikel:" + ids
        shop_ws.cell(row=total_rows, column=8).value = "Circ IT"
        shop_ws.cell(row=total_rows, column=9).value = "20"
        shop_ws.cell(row=total_rows, column=12).value = "Stück"
        shop_ws.cell(row=total_rows, column=16).value = bundle.sum
        shop_ws.cell(row=total_rows, column=20).value = "EUR"
        shop_ws.cell(row=total_rows, column=21).value = "19"
        shop_ws.cell(row=total_rows, column=21).value = str(new_id) + ".png"
        shop_ws.cell(row=total_rows, column=25).value = "png"

        shop_ws.cell(row=total_rows, column=16).number_format = '#,##0.00€'
        shop_ws.cell(row=total_rows, column=21).number_format = "00"

        total_rows += 1

        new_id = increment_id(new_id)

    shop_wb.save("shop_list.xlsx")

    return gr.update(value=path.abspath("shop_list.xlsx"))


with gr.Blocks(theme=gr.themes.Soft()) as app:
    with gr.Row():
        price_list = gr.File(label="Preisliste")
        shop_list = gr.File(label="Shopliste", value="shop_template.xlsx", visible=False)

    send_button = gr.Button("Übertragen")

    output = gr.File(label="Neue Shopliste")

    send_button.click(transfer_data, inputs=[price_list, shop_list], outputs=[output])

app.queue()
app.launch(show_error=True, server_port=8080, ssl_verify=True)
