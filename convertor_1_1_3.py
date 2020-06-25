import argparse
import csv
import glob
import logging
import os
import shutil
import smtplib
import tkinter as tk
from csv import DictWriter
from datetime import datetime
import time
from os.path import basename, splitext, join
from tkinter import filedialog, simpledialog
from tkinter.scrolledtext import ScrolledText
from bs4 import BeautifulSoup
from dxf2svg.pycore import save_svg_from_dxf
from reportlab.graphics import renderPDF
from svglib.svglib import svg2rlg


class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""

    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(tk.END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(tk.END)

        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


class GUI(tk.Frame):
    """ This class defines the graphical user interface """

    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.root = parent
        self.generate_button = tk.Button(self.root, text="КОНВЕРТИРАНЕ ОТ ДИРЕКТОРИЯ", command=main)
        self.goto_button = tk.Button(self.root, text='ОТВОРИ ГОТОВИ ЗАЯВКИ', command=gotodir)
        self.report_error = tk.Button(self.root, text="ДОКЛАДВАЙ ПРОБЛЕМ", command=report_error)
        self.text_handler = None
        self.build_gui()

    def build_gui(self):
        self.root.title('GlassOrderConvertor')
        self.root.wm_iconbitmap('logo.ico')
        self.generate_button.grid(row=0, column=0)
        self.goto_button.grid(row=0, column=1)

        # Add ScrolledText widget to display logging
        st = ScrolledText()
        st.configure(font='Areal', width=150)
        st.grid(row=1, sticky='ew', columnspan=2)

        # Add button report_error
        self.report_error.grid(row=3, columnspan=2)

        # Create textlogger
        self.text_handler = TextHandler(st)


################################################################################
# Main Program definitions
#
def makesvg(drawings, dirpath, msg=None):
    for d in drawings:
            draw_file = d.split(',')
            dxf_input = (str(draw_file).replace("['", "").replace("']", ""))
            svg_output = (
                str(draw_file).replace("['", "").replace("']", "").replace(".dxf", ".svg"))
            try:
                save_svg_from_dxf(str(dxf_input), size=6000)
                try:
                    shutil.move(dxf_input, dirpath + '/success/')
                except:
                    msg = str(dxf_input).replace(dirpath, '')
                    logger.exception("Има проблем с файл %s", msg)
            except:
                msg = str(draw_file).replace(dirpath, '')
                logger.exception("Има проблем с файл %s", msg)
            try:
                shutil.move(svg_output, dirpath + "/GlassPurchaseOrders/drawing_pdf")
            except:
                msg = str(svg_output).replace(dirpath, '')
                logger.exception("Има проблем с файл %s", msg)
            return msg


def makepdf(dirpath):
    svg_file = glob.glob(dirpath + "/GlassPurchaseOrders/drawing_pdf/*.svg")
    if len(svg_file) != 0:
        drawingsvg = svg_file
        for d in drawingsvg:
            draw_file_svg = d.split(',')
            try:
                svg_output = (str(draw_file_svg).replace("['", "").replace("']", ""))
                pdf_output = (
                    str(draw_file_svg).replace("['", "").replace("']", "").replace(".svg", ".pdf"))
                drawing = svg2rlg(str(svg_output))
                renderPDF.drawToFile(
                    drawing.resized(kind='fit', lpad=150, rpad=150, bpad=150, tpad=150),
                    str(pdf_output))
                try:
                    os.remove(svg_output)
                except:
                    logger.exception(
                        (str(svg_output).replace(dirpath, '')) + ' Не може да бъде премахнат от директорията')
                    pass

            except:
                logger.exception(str(draw_file_svg).replace(dirpath, '') + " Не може да бъде конвертиран в PDF")
                pass
    else:
        pass


def soupCooking(member, dirpath, xfiles, csv_output):
    try:
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
        istherescetch = member.findNext('dxfsketch').isSelfClosing
        if arch == 0 and spros == 0 and istherescetch is True:
            pass
        else:
            docnum = member.findPrevious('document_number')
            logger.info(
                "Ще бъдат конвертирани допълнителни чертежи към заявка:" + docnum.text + ", моля проверете в PDF директория")
            drawings = glob.glob(dirpath + "/*.dxf")
            if len(drawings) > 0:
                makesvg(drawings, dirpath)
                makepdf(dirpath)
            else:
                logger.error(str(scetch.text + ".dxf") + ' не бе намерен в целевата директория')
    except AttributeError:
        logger.error("Проверете xml файлa за липсващи атрибути")
        logger.exception("Липсващи атрибути, %s",
                             'Проверете xml файл ' + str(xfiles).replace(dirpath, ''))
        pass

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
            mem['archinfo'] = "Scetch No.:" + str(scetch.text) + " Radius = " + str(float(radius.text) / 1000) + " " + "Rise = " + str(float(rise.text) / 1000) + " " + "X = " + str(float(x.text) / 1000) + " " + "Y = " + str(float(y.text) / 1000)
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
        row = {'ITEM': mem['item'], 'DESCRIPTION': mem['desc'], 'HEIGHT': mem['height'], 'WIDTH': mem['width'],
                   'ALU_spacer': mem['spacer'], 'BARCODE': mem['barcode'], 'FIELD': mem['field'],
                   'Special_info': mem['archinfo'] + ' ' + mem['sprosinfo']}
        csv_output.writerow(row)


def loadandconvert(xml_file, dirpath):
    filename = xml_file
    for f in filename:
        file = f
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
        xfiles = str(file).replace("['", "").replace("']", "")
        cfiles = str(file).replace("['", "").replace("']", "").replace(".xml", ".csv")
        logger.info('Отварям файл: ')
        try:
            with open(xfiles, encoding='UTF-8') as f_input, open(str(cfiles).replace(dirpath, dirpath + "/GlassPurchaseOrders/mawi_csv/"), 'w',encoding='UTF-8',newline='') as f_output:
                csv_output: DictWriter = csv.DictWriter(f_output, fieldnames=fieldnames, dialect='unix')
                csv_output.writeheader()

                xml = f_input.read()
                soup = BeautifulSoup(xml, 'xml')
                for member in soup.findChildren('barcode_l'):
                   soupCooking(member, dirpath, xfiles, csv_output)


        except:
            logger.exception('Грешка при обработка на файл / %s',str(xfiles).replace(dirpath, '') + 'моля опитайте отново')
            pass
        sortfiles(file, dirpath, xfiles, cfiles)

def sortfiles(file, dirpath, xfiles, cfiles):
    try:
        with open(str(file).replace("['", "").replace("']", ""), encoding='UTF-8') as f_input:
            xml = f_input.read()
            soup = BeautifulSoup(xml, 'xml')
            deliveryinfo = soup.find('order_remark_mawi')

        try:
            try:
                os.rename(str(cfiles).replace(dirpath, dirpath + "/GlassPurchaseOrders/mawi_csv/"),
                          str(cfiles).replace(dirpath, dirpath + "/GlassPurchaseOrders/mawi_csv/").replace(".csv",
                                                                                                          " - " + str(
                                                                                                              deliveryinfo.text) + ".csv"))
                logger.info("Документ:" + str(cfiles).replace(dirpath, '') + " бе конвертиран успешно")
            except FileExistsError:
                logger.info("Документ:" + str(cfiles).replace(dirpath,'') + 'вече съществува в целевата директория')
                pass
            except:
                logger.info("Документ:" + str(cfiles).replace(dirpath,
                                                              '') + " не може да се да се добави:" + deliveryinfo.text + " към името на файла, моля преименувайте ръчно")
                pass
            shutil.move(xfiles, dirpath + '/success/')
            logger.info("Документ:" + str(xfiles).replace(dirpath, '') + " бе преместен в папка 'success'")
        except FileExistsError:
            logger.info(
                str(xfiles).replace(dirpath, '') + ' File has been created before, and therefore it will be deleted')
            pass
        except FileNotFoundError:
            logger.info(str(xfiles).replace(dirpath, '') + ' File has not been found')
            pass
    except FileExistsError:
        logger.info(
            str(xfiles).replace(dirpath, '') + ' File has been created before, and therefore it will be deleted')
        os.remove(xfiles)
        pass
    except FileNotFoundError:
        logger.info(str(xfiles).replace(dirpath, '') + ' File has not been found')
        pass
    except shutil.Error:
        logger.info(str(cfiles).replace(dirpath, '') + " Файлът вече съществува в директорията")
        logging.info(cfiles + ' ' + str(shutil.Error))
        os.remove(xfiles)
        os.remove(cfiles)
        pass


def gotodir(dirpath):
    os.system("start" + dirpath + "\GlassPurchaseOrders")


def deloldlog(days=0):
    file = glob.glob('log/*.log')
    for f in file:
        file_time = os.path.getmtime(f)
        if (time.time() - file_time) / 3600 > 24 * days:
            os.remove(f)
        else:
            pass


################################################################################
# Main Program Flow
#
def main():
    dirpath = str(tk.filedialog.askdirectory()).replace("['", "").replace("']", "")
    xml_file = glob.glob(dirpath + "/*.xml")
    if len(xml_file) != 0:
        try:
            loadandconvert(xml_file, dirpath)
            try:
                deloldlog(1)
            except:
                logger.exception("Грешка %s", "Не могат да се изтрият старите логовее")
        except:
            logger.exception("Настъпи неочаквана грешка %s", "Моля проверете заявките")
            pass
    else:
        logger.info("Не бяха подготвени заявки / %s", "липсват валидни xml файлове в директорията")
        pass


################################################################################
# Read commandline arguments
#
def get_arguments():
    parser = argparse.ArgumentParser(
        description="Test GUI logging")
    parser.add_argument('--logdir', required=False, default='./log/')
    parser.add_argument('--debug', action="store_const", const=logging.DEBUG, default=logging.INFO)
    a = parser.parse_args()
    return a


################################################################################
# Configure logger
#
def get_logger(log_level=logging.INFO, log_dir=None, text_handler=None):
    script = splitext(basename('GOC.py'))[0]
    logg = logging.getLogger(script)
    logg.setLevel(log_level)

    # set up file or stdout handlers
    if log_dir:
        info_file = join(log_dir, script + str(datetime.now()).replace(':', '_') + '.log')
        info_handler = logging.FileHandler(info_file, encoding='utf-8')
    else:
        info_handler = logging.StreamHandler()

    # create formatter and add it to the handlers
    formatter = logging.Formatter(u"%(asctime)s [%(levelname)s] %(message)s")
    info_handler.setFormatter(formatter)
    info_handler.setLevel(log_level)
    info_handler.encoding = 'utf-8'
    if text_handler:
        text_handler.setFormatter(formatter)
        text_handler.setLevel(log_level)
        text_handler.encoding = 'utf-8'

    # add the handlers to logg
    logg.addHandler(info_handler)
    if text_handler:
        logg.addHandler(text_handler)

    return logg


################################################################################
# Connecting to an SMTP Server

#
def report_error():
    textfile = str(logger.handlers)[14:-30]
    sender = os.getlogin() + '@teolino.eu'
    password = tk.simpledialog.askstring('password', 'Въведете парола за ' + sender, show='*')
    helpdesk = 'ivan.goranov@teolino.eu'
    # Import the email modules we'll need
    from email.mime.text import MIMEText


    fp = open(str(textfile).replace('<Logger ', './').replace(' (INFO)>', '.log'), 'r', encoding='utf-8')
    # Create a text/plain message
    msg = MIMEText(fp.read(), _charset='utf-8')
    fp.close()

    msg['Subject'] = 'GlassConvertorError at %s' % textfile
    msg['From'] = sender
    msg['To'] = helpdesk

    # Send the message
    try:
        s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        s.ehlo()
        s.login(sender, password)
        s.sendmail(sender, [helpdesk], msg.as_string())
        s.quit()
    except UnicodeEncodeError:
        logger.info('моля сменете езика си на английски')
    except:
        logger.exception(
            'Неуспешен опит за изпращане на мейл, моля копирайте текста с CTR+c и го пратете ръчно до: ivan.goranov@teolino.eu')


################################################################################
# get args, configure logger and launch GUI
#
if __name__ == '__main__':
    args = get_arguments()
    root = tk.Tk()
    gui = GUI(root)
    logger = get_logger(log_level=args.debug, log_dir=args.logdir, text_handler=gui.text_handler)
    root.mainloop()
