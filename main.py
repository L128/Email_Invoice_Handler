# -*- coding: utf-8 -*-
import configparser
import os
import time
import tempfile
import pdfplumber
import ssl
import pyMail

ssl._create_default_https_context = ssl._create_unverified_context


# import requests
# from bs4 import BeautifulSoup
# from urllib.request import Request, urlopen
# from fake_useragent import UserAgent


def get_data(directory):
    config = configparser.ConfigParser()
    config.read(directory)
    return [config["email"]["Imap_server"],
            config["email"]["Smtp_server"],
            config["email"]["Email_address"],
            config["email"]["Password"],
            config["email"]["Target"],
            config["email"]["Archive_folder"],
            config["invoice"]["Company_name"],
            config["invoice"]["TaxpayerID"]]


def if_pdf_invoice(pdf, company_name, taxpayerID):
    text = pdf.pages[0].extract_text()
    if text.count("电子普通发票") & \
            text.count(company_name) & \
            text.count(taxpayerID) & \
            text.count("发票号码") & \
            text.count("开票日期") & \
            text.count("(小写)"):
        return text[text.index("发票号码") + 6:text.index("发票号码") + 13]
    else:
        return False


def pull_pdf_data(pdf):
    text = pdf.pages[0].extract_text()
    try:
        invoice_num = text[text.index("发票号码") + 6:text.index("发票号码") + 13]
    except Exception as e:
        invoice_num = '无法识别'

    try:
        text = text[text.find("开票日期") + 5:]
        invoice_date = text[:text.index("日")+1]
    except Exception as e:
        invoice_date = '无法识别'

    try:
        text = text[text.find("(小写)") + 4:]
        total_price = text[:text.index(" ", text.index("名"))-3]
    except Exception as e:
        text = '无法识别'
    line = '发票号码： ' + invoice_num + '\n开票日期： ' + invoice_date.replace(' ', '') + \
           '\n总金额： ' + total_price + '\n------------------------------------------\n'
    return line


def cache_invoices_attachments(attachment_list, tempdir):
    if attachment_list['name'][-4:] == '.pdf':
        f_name = os.path.join(tempdir, 'tmp.pdf')
        with open(f_name, 'wb') as f:
            f.write(attachment_list['data'])
            f.close()
            invoice_num = if_pdf_invoice(pdfplumber.open(f_name), Company_name, TaxpayerID)
            if invoice_num:
                # rename the file with invoice number
                os.rename(f_name, os.path.join(tempdir, invoice_num + '.pdf'))
                return True
    else:
        return False


def check_invoice_attachments(attachment_list, tempdir):
    at_least_one_invoice = False
    if attachment_list:
        for attachment in attachment_list:
            at_least_one_invoice = at_least_one_invoice or cache_invoices_attachments(attachment, tempdir)
    return at_least_one_invoice


def check_invoice_links(links): pass


# Can't do it at the moment. Will add it later.

# if links:
#     for link in links:
#         print(link)
#         try:
#             # data = urllib.request.urlopen(link).read()
#             data = requests.get(link)
#             bs = BeautifulSoup(data.text(), "html.parser")
#             for i in BeautifulSoup.find_all('a', {'class': "fpdown"}):
#                 print(re.search('.*\.pdf', i.get('href')).group(0))
#         except Exception as e:
#             print(e)


def send_email_to_target(target, tempdir): pass


def archive_an_email(email_uid, des):
    result = rml.mail.uid('COPY', email_uid, des)
    if result[0] == 'OK':
        rml.mail.uid('STORE', email_uid, '+FLAGS', '(\Deleted)')
        print("Archived")
        return rml.mail.expunge()
    else:
        print("Not Archived")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    [Imap_server, Smtp_server, Email_address, Password, Target, Archive_folder, Company_name, TaxpayerID] \
        = get_data("./credentials.conf")
    rml = pyMail.ReceiveMailDealer(Imap_server, Email_address, Password)
    sml = pyMail.SendMailDealer(Smtp_server, Email_address, Password)

    if rml.getInbox()[1][0].split() != '':
        email_title = Company_name + time.strftime('%Y%m') + "发票"
        email_body = ""
        email_attachment_dir = []
        with tempfile.TemporaryDirectory() as td:
            for num in rml.getInbox()[1][0].split():
                mailInfo = rml.getMailInfo(num)
                if check_invoice_attachments(mailInfo['attachments'], td):
                    archive_an_email(rml.getEmailUID(num), Archive_folder)
                # else:
                #     check_invoice_links(mailInfo['html'])

            if os.path.exists(os.path.join(td, 'tmp.pdf')):
                os.remove(os.path.join(td, 'tmp.pdf'))
            # Pulling attachment table from cache
            if os.listdir(td):
                for invoice in os.listdir(td):
                    if invoice.endswith(".pdf"):
                        email_body += pull_pdf_data(pdfplumber.open(os.path.join(td, invoice)))
                        email_attachment_dir.append(os.path.join(td, invoice))
                print(email_title)
                print(email_body)
                # print(email_attachment_dir)
                sml.setMailInfo(Target, email_title, email_body, 'plain', email_attachment_dir)
                if input(f"是否确认以上信息并发送给 {Target} 确认请输入'Y'： ") == 'Y':
                    sml.sendMail()
                    print("已发送")
