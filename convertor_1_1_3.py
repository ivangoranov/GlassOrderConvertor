import argparse
import glob
import logging
import os
import smtplib
import tkinter as tk
from datetime import datetime
from os.path import basename, splitext, join
from tkinter import filedialog, simpledialog
from tkinter.scrolledtext import ScrolledText

import MainProgramDefinitions
from MainProgramDefinitions import soup_cooking, sortfiles

global dirpath, file, finaldirpath

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
# Main Program Flow
#
def main():
    dirpath = str(tk.filedialog.askdirectory()).replace("['", "").replace("']", "")
    xml_file = glob.glob(dirpath + "/*.xml")
    for x in xml_file:
        file = x
        logger.info('Отварям файл:  %s', str(file).replace(dirpath, ""))
        try:
            soup_cooking(file, dirpath)
        except:
            logger.exception("Настъпи неочаквана грешка %s", "Моля проверете заявките")
            pass
        try:
            sortfiles(file, dirpath)
            logger.info("Документ:" + str(file).replace(dirpath, '') + " бе конвертиран успешно")
        except:
            logger.exception('Грешка при местене на файл / %s', str(file).replace(dirpath, '') + 'моля опитайте отново')
            pass



def gotodir():
    os.system("start " + str(finaldirpath) + "\GlassPurchaseOrders")


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
