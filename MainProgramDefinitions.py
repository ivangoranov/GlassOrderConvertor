################################################################################
# Main Program definitions
#
import datetime
import glob
import csv
import os
import shutil
import time
from datetime import datetime
from csv import DictWriter
from bs4 import BeautifulSoup
from dxf2svg.pycore import save_svg_from_dxf
from logger import logger
from reportlab.graphics import renderPDF
from svglib.svglib import svg2rlg


def makesvg(drawings, dirpath):
    for d in drawings:
        draw_file = d.split(',')
        dxf_input = (str(draw_file).replace("['", "").replace("']", ""))
        svg_output = (str(draw_file).replace("['", "").replace("']", "").replace(".dxf", ".svg"))
        save_svg_from_dxf(str(dxf_input), size=10000)
        os.rename(dxf_input, str(dxf_input).replace(dirpath, dirpath + '/success/'))
        os.rename(svg_output, str(svg_output).replace(dirpath, dirpath + "/GlassPurchaseOrders/drawing_pdf"))


def makepdf(dirpath):
    svg_file = glob.glob(dirpath + "/GlassPurchaseOrders/drawing_pdf/*.svg")
    drawingsvg = svg_file
    for d in drawingsvg:
        draw_file_svg = d
        svg_output = (str(draw_file_svg).replace("['", "").replace("']", ""))
        pdf_output = (str(draw_file_svg).replace("['", "").replace("']", "").replace(".svg", ".pdf"))
        drawing = svg2rlg(str(svg_output))
        renderPDF.drawToFile( drawing.resized(kind='expandx', lpad=150, rpad=150, bpad=150, tpad=150),str(pdf_output))
        os.remove(svg_output)


def soup_cooking(file, dirpath, dxf=False):
    global item, field, clap, desc, glass_hight, glass_width, spacer, arch, spros, scetch, radius, rise, x, y, sbkey, sbdesc, instkind
    fields = [
        "ITEM",
        "FIELD",
        "DESCRIPTION",
        "WIDTH",
        "HEIGHT",
        "ALU_spacer",
        "Special_info",
        "BARCODE"
    ]

    fieldnames = fields
    cfiles = str(file).replace(".xml", ".csv")
    with open(file, encoding='UTF-8') as f_input, open(str(cfiles), 'w', encoding='UTF-8', newline='') as f_output:
        csv_output: DictWriter = csv.DictWriter(f_output, fieldnames=fieldnames, dialect='unix')
        csv_output.writeheader()
        xml = f_input.read()
        soup = BeautifulSoup(xml, 'xml')
        for member in soup.findChildren('barcode_l'):
            mem = member
            instkind = member.findPrevious('deliverykind')
            item = member.findPrevious('item_number')
            field = member.findPrevious('field_nr')
            desc = member.findPrevious('product_des')
            glass_hight = member.findNext('glassheight')
            glass_width = member.findNext('glasswidth')
            spacer = member.findNext('spacer').key
            if member.findNext('archinformation').isSelfClosing:
                arch = 0
                scetch = ''
            else:
                arch = 1
                scetch = member.findPrevious('sketch')
                x = member.findNext('archdata').x_dim
                y = member.findNext('archdata').y_dim
                radius = member.findNext('archdata').radius
                rise = member.findNext('archdata').rise
            if member.findNext('glass_sashbar').isSelfClosing:
                spros = 0
                if scetch != '':
                    scetch = scetch
                else:
                    pass
            else:
                spros = 1
                scetch = member.findPrevious('sketch')
                sbkey = member.findNext('sashbardata').sashbar_key
                sbdesc = member.findNext('sashbardata').sashbar_text
            clap = member.findNext('pressure_balance')
            if arch == 0 and spros == 0:
                return dxf
                pass
            else:
                docnum = member.findPrevious('document_number')
                logger.info(
                    "Ще бъдат конвертирани допълнителни чертежи към заявка:" + docnum.text + ", моля проверете в PDF директория")
                drawings = glob.glob(dirpath + "/" + str(docnum.text) + "*.dxf")
                if len(drawings) > 0:
                    makesvg(drawings, dirpath)
                    makepdf(dirpath)
                    return dxf is True
                else:
                    logger.error(str(scetch.text + ".dxf") + ' не бе намерен в целевата директория')
            mem['item'] = item.text
            mem['field'] = field.text
            if clap is None or clap.text == 'false':
                mem['desc'] = desc.text
            else:
                mem['desc'] = desc.text + ' С КЛАПАН'
            mem['height'] = glass_hight.text
            mem['width'] = glass_width.text
            if spacer is None:
                mem['spacer'] = ""
            else:
                mem['spacer'] = str(spacer.text).strip()
            if arch == 1 and spros == 0:
                mem['archinfo'] = 'Scetch No.:' + str(scetch.text) + " Radius = " + str(
                    float(radius.text) / 1000) + " " + "Rise = " + str(
                    float(rise.text) / 1000) + " " + "X = " + str(float(x.text) / 1000) + " " + "Y = " + str(
                    float(y.text) / 1000)
                mem['sprosinfo'] = ''
            elif arch == 0 and spros == 1:
                mem['archinfo'] = ''
                mem['sprosinfo'] = "Scetch No.:" + str(
                    scetch.text) + ' / ' + sbkey.text + ' - ' + sbdesc.text
            elif arch == 1 and spros == 1:
                mem['archinfo'] = "Scetch No.:" + str(scetch.text) + " Radius = " + str(
                    float(radius.text) / 1000) + " " + "Rise = " + str(
                    float(rise.text) / 1000) + " " + "X = " + str(
                    float(x.text) / 1000) + " " + "Y = " + str(
                    float(y.text) / 1000)
                mem['sprosinfo'] = sbkey.text + ' - ' + sbdesc.text
            elif arch == 0 and spros == 0:
                mem['archinfo'] = ''
                mem['sprosinfo'] = ''
            if instkind.text == 'MONTAGEART_INFERTIGUNG':
                mem['barcode'] = mem.text
            else:
                mem['barcode'] = ''
            row = {'ITEM': mem['item'], 'DESCRIPTION': mem['desc'], 'HEIGHT': mem['height'],
                   'WIDTH': mem['width'],
                   'ALU_spacer': mem['spacer'], 'BARCODE': mem['barcode'], 'FIELD': mem['field'],
                   'Special_info': mem['archinfo'] + ' ' + mem['sprosinfo']}
            csv_output.writerow(row)


def sortfiles(file, dirpath):
    with open(str(file), encoding='UTF-8') as f_input:
        cfile = str(file).replace('.xml', '.csv')
        xml = f_input.read()
        soup = BeautifulSoup(xml, 'xml')
        deliveryinfo = soup.find('order_remark_mawi')
        deliverydate = soup.find('glasses_delivery_date_mawi').text
        deliverydate_obj = datetime.strptime(deliverydate, '%Y-%m-%d')
        deliverydatestr = deliverydate_obj.strftime('%d.%m.%Y')
        mawiorderno = str(cfile)[-13:]
        os.rename(cfile, (dirpath + '/GlassPurchaseOrders/mawi_csv/' + "Order_" + str(deliveryinfo.text) + "_" + str(
            deliverydatestr) + str(mawiorderno)))
    os.rename(file, str(file).replace(dirpath, dirpath + '/success/'))


def deloldlog(days=0):
    file = glob.glob('/log/*.log')
    for f in file:
        file_time = os.path.getmtime(f)
        if (time.time() - file_time) / 3600 > 24 * days:
            os.remove(f)
        else:
            pass
