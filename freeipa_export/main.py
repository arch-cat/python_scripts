# -*- coding: utf-8 -*-
# Imports
import warnings
import os
import logging
import paramiko
import sys
import colorama
from python_freeipa import ClientMeta
from datetime import datetime
from csv import writer

import win32com.client as win32

import settings
import logger

warnings.filterwarnings("ignore")
colorama.init()


def export_csv_data():
    current_date = datetime.now().strftime("%d-%m-%Y")

    user = ""
    port = ""
    database = ""

    sql_query = f"SELECT * " f"FROM "

    command = (
        f"psql -h  "
        f"-U {user}@ "
        f"-d {database} "
        f"-p {port} "
        f'-c "\\copy ({sql_query}) '
        f"to '/home/users.csv' delimiter ';' csv header;\""
    )

    with paramiko.SSHClient() as client:
        _start_time = datetime.now()

        client.load_system_host_keys()
        logging.info("SSH connection")
        client.connect(
            hostname=settings.SSH_HOST,
            username=settings.SSH_USER,
            password=settings.SSH_PASSWORD,
        )
        logging.info("Successfully logged in as ")
        logging.info("Execute SQL query:\n" f"{sql_query}")
        logging.info("Save to remote host, path: /home/users.csv")
        (stdin, stdout, stderr) = client.exec_command(command)
        output = stdout.readlines()
        for line in output:
            logging.info(f"Query success: {line.rstrip()}")

        logging.info("Download file /home/users.csv")
        remote_path = "/home/users.csv"
        sftp_client = client.open_sftp()
        sftp_client.get(remote_path, "./csv/sudir.csv")
        logging.info("Download success")

        logging.info('Change encoding file to "utf-8-sig"')
        with open("./csv/sudir.csv", encoding="utf-8") as before_file:
            data = before_file.read()
        with open("./csv/sudir.csv", "wb") as after_file:
            after_file.write(data.encode("utf-8-sig"))
        try:
            os.rename("./csv/sudir.csv", f"./csv/SUDIR_{current_date}.csv")
        except FileExistsError:
            os.remove(f"./csv/SUDIR_{current_date}.csv")
            os.rename("./csv/sudir.csv", f"./csv/SUDIR_{current_date}.csv")
        logging.info("Change encoding file success")
        return _start_time


def open_txt():
    with open("group.txt", "r") as file_groups:
        _groups = file_groups.read().splitlines()
    return _groups


def add_data_to_csv(_start_time):
    current_date = datetime.now().strftime("%d-%m-%Y")

    logging.info("Connect to IPA")
    client = ClientMeta(settings.HOST_IPA, verify_ssl=False)
    client.login(settings.USER_IPA, settings.PASSWORD_IPA)
    _groups = open_txt()

    _keys = [
        "uid",
        "displayname",
        "ipauniqueid",
        "employeenumber",
        "telephonenumber",
        "nsaccountlock",
        "mail",
        "memberof_group",
    ]

    _exceptions_count = 0
    _user_count = 0
    _csv_data_count = 0
    data = []
    with open(f"./csv/IPA_{current_date}.csv", "w", newline="") as f_object:
        for i_group in _groups:
            group = client.group_show(f"{i_group}")
            users = group.get("result")["member_user"]
            for i in users:
                if i[0] == "u" and i[1] == "_" in i:
                    logging.warning(f"[EXCEPTION] ['{i}']")
                    _exceptions_count += 1
                    continue
                else:
                    ipa_user = client.user_show(i)
                    _user_count += 1
                    for key in _keys:
                        if key == "memberof_group":
                            string = f"['{i_group}']"
                        else:
                            try:
                                string = f"{ipa_user.get('result')[key]}"
                            except KeyError:
                                string = "['None']"
                        for ch in ["[", "]", "'"]:
                            if ch in string:
                                string = string.replace(ch, "")
                        data.append(string)
                logging.info(f"[ADD_TO_CSV] {data}")
                writer_object = writer(f_object, delimiter=";")
                writer_object.writerow(data)
                data.clear()
    f_object.close()
    _end_time = datetime.now() - _start_time
    _end_time = str(_end_time).split(".")[0]

    client.logout()
    return _exceptions_count, _user_count, _end_time


def send_msg(_exceptions, _user_count, _end_time):
    current_date = datetime.now().strftime("%d-%m-%Y")

    _groups = open_txt()
    nl = "<br>"
    csv_file_ipa = f"./csv/IPA_{current_date}.csv"
    csv_file_sudir = f"./csv/SUDIR_{current_date}.csv"
    log_file = f"./logs/export_users_{current_date}.log"

    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = ""  # Email destination "Email#1;Email#2"
    mail.Subject = f"Выгрузка из IPA_{current_date} & SUDIR_{current_date}"
    mail.HTMLBody = (
        f"<p>Сформировано автоматически</p>"
        f"<p>Список групп:<br>"
        f"{nl.join(_groups)}</p>"
        f"Количество пользователей во всех группах: {_user_count}<br>"
        f"Количество исключений: {_exceptions}<br>"
        f"Время выполнения {_end_time}</p>"
        f"<p>-----------------------------------------------------<br>"
        f"(C) 2023 Grif Ivan, Moscow, Russia<br>"
        f"Released under GNU Public License (GPL)<br>"
        f"<a href=https://bitbucket.ru/projects/repos/"
        f"browse/support/grif_iv?at=support>Bitbucket</a></p>"
    )
    mail.Attachments.Add(os.path.join(os.getcwd(), csv_file_ipa))
    mail.Attachments.Add(os.path.join(os.getcwd(), csv_file_sudir))
    mail.Attachments.Add(os.path.join(os.getcwd(), log_file))
    mail.Send()


def main():
    current_date = datetime.now().strftime("%d-%m-%Y")

    script_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    days = (
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    )
    now_dt = datetime.now()
    current_weekday = days[now_dt.weekday()]
    if not logger.set_up_logging(
        console_log_output="stdout",
        console_log_level="info",
        console_log_color=True,
        logfile_file=f"./logs/{script_name}_{current_date}.log",
        logfile_log_level="debug",
        logfile_log_color=False,
        log_line_template=f"%(color_on)s[{current_weekday} %(asctime)s] [%(threadName)s] [%(levelname)s] :::: %("
        f"message)s%(color_off)s",
    ):
        print("Failed to set up logging, aborting.")
        return 1

    logging.info(f"Begin {script_name} script")

    start_time = export_csv_data()

    exceptions_count, user_count, end_time = add_data_to_csv(start_time)
    logging.info(f"End {script_name} script")
    logging.info(f"Lead time: {end_time}")
    logging.info(f"User count: {user_count}")
    logging.info(f"Exceptions count: {exceptions_count}")

    send_msg(exceptions_count, user_count, end_time)


if __name__ == "__main__":
    sys.exit(main())
