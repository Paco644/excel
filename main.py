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


def send(price_list, shop_list, mode, settings):
    if mode == 1:
        return send_mode_init(price_list, shop_list)
    else:
        raise gr.Error("Noch nicht implementiert.")


def send_mode_init(price_list, shop_list):
    if price_list is None:
        raise gr.Error("Bitte die Preisliste hochladen.")

    if shop_list is None:
        raise gr.Error("Bitte die Shopliste hochladen.")

    shutil.copy(price_list.name, "price_list.xlsx")

    price_wb: Workbook = load_workbook("price_list.xlsx")
    shop_wb: Workbook = load_workbook("shop_template.xlsx")

    try:
        price_ws: Worksheet = price_wb['40495396']
        shop_ws: Worksheet = shop_wb['Artikel']
    except KeyError:
        raise gr.Error("Preisliste enthält falsche daten.")

    price_data = price_ws['A:D']
    name = price_data[0]
    price = price_data[3]
    bundle = None
    new_id = increment_id("HW-CIT-0000")
    total_rows = 4

    bundles: list[Bundle] = []
    products: list[Product] = []

    for i in range(1, len(name) - 5):
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
    for bundle in bundles:
        products += bundle.products

    # Alle Duplikate entfernen
    scanned = []
    for product in products:
        if product.name not in scanned:
            scanned.append(product.name)
        else:
            products.remove(product)

    # IDS für die Einzelrprodukte setzen
    for product in products:
        product.id = new_id
        new_id = increment_id(new_id)

    # Einzelprodukte in die Excel hinzufügen
    for product in products:
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
    for bundle in bundles:

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

    return gr.update(value=path.abspath("shop_list.xlsx"), visible=True)


modes_value = ["Zusammenführen (Nicht implementiert)", "Initiation"]
default_mode = modes_value[0]

settings_value = ["Zeilen beim Zusammenführen aktualisieren"]
default_settings = [settings_value[0]]


def on_mode_changed(mode):
    if mode == 0:
        return [gr.update(visible=False, value=None), gr.update(visible=False, value="shop_template.xlsx")]
    else:
        return [gr.update(visible=True, value=default_settings), gr.update(visible=True, value=None)]


with gr.Blocks() as app:
    with gr.Row():
        mode = gr.Radio(label="Modus", choices=modes_value, value=default_mode, interactive=True,
                        type="index")
        settings = gr.CheckboxGroup(choices=settings_value, value=default_settings, label="Einstellungen (Vorschau)",
                                    type="index", interactive=True)
    with gr.Row():
        price_list = gr.File(label="Preisliste")
        shop_list = gr.File(label="Shopliste")

    info = gr.Label(visible=False)
    send_button = gr.Button("Übertragen")

    output = gr.File(label="Neue Shopliste", visible=False)

    send_button.click(send, inputs=[price_list, shop_list, mode, settings], outputs=[output])

    mode.select(on_mode_changed, inputs=mode, outputs=[settings, shop_list])

app.queue()
app.launch(show_error=True, server_port=8080, ssl_verify=True)
