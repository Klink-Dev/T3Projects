import io
import json
from time import sleep

import urllib3
import requests
import pendulum
import pandas as pd
from tabulate import tabulate
from urllib3.util import Retry
from requests.adapters import HTTPAdapter
from urllib3.exceptions import MaxRetryError
from requests_kerberos import OPTIONAL, HTTPKerberosAuth
from apscheduler.schedulers.blocking import BlockingScheduler


def timezone(fc: str) -> str:
    url = f"https://infrascripts.amazon.com/get_site_details_json/{fc}"
    with requests_retry_session() as req:
        resp = req.get(url)
        try:
            data = json.loads(resp.text)["tbl_site_data"]["timeZone"]
            return data
        except TypeError:
            print(f"Time Zone not found for {fc} in API!")
            return None


def requests_retry_session(retries=5,
                           session=None,
                           backoff_factor=0.3,
                           status_forcelist=(400, 500, 502, 503, 504, 520, 524),
                           method_whitelist=("HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS", "TRACE")):
    urllib3.disable_warnings()

    try:
        session = session or requests.Session()
        session.verify = False
        session.auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL)

        retry = Retry(read=retries,
                      total=retries,
                      connect=retries,
                      backoff_factor=backoff_factor,
                      method_whitelist=method_whitelist,
                      status_forcelist=status_forcelist)

        adapter = HTTPAdapter(max_retries=retry)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        return session

    except MaxRetryError:
        print("error")
        pass


def get_fc_roster(fc: str) -> pd.DataFrame:
    url = (f"https://fclm-portal.amazon.com/employee/employeeRoster?reportFormat=CSV&warehouseId={fc}&"
           f"employeeStatusActive=true&_employeeStatusActive=on&employeeStatusTerminated=true&_"
           f"employeeStatusTerminated=on&employeeStatusLeaveOfAbsence=true&_employeeStatusLeaveOfAbsence=on&"
           f"employeeStatusExempt=true&_employeeStatusExempt=on&employeeTypeAmzn=true&_employeeTypeAmzn=on&"
           f"employeeTypeTemp=true&_employeeTypeTemp=on&employeeType3Pty=true&_employeeType3Pty=on&"
           f"Employee+ID=Employee+ID&User+ID=User+ID&Employee+Name=Employee+Name&Badge+Barcode+ID=Badge+Barcode+ID&"
           f"Department+ID=Department+ID&Employment+Start+Date=Employment+Start+Date&Employment+Type="
           f"Employment+Type&Employee+Status=Employee+Status&Manager+Name=Manager+Name&Temp+Agency+Code="
           f"Temp+Agency+Code&Job+Title=Job+Title&Management+Area+ID=Management+Area+ID&Shift+Pattern="
           f"Shift+Pattern&Badge+RFID=Badge+RFID&Exempt=Exempt&hideColumns=Photo&submit=true")

    with requests_retry_session() as req:
        resp = req.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL), verify=False,
                       allow_redirects=True, timeout=30)

    roster_df = pd.read_csv(io.StringIO(resp.text))

    return roster_df


def get_ppr(fc: str, start_date: pendulum, end_date: pendulum):
    url = (f"https://fclm-portal.amazon.com/reports/processPathRollup?reportFormat=CSV&warehouseId={fc}"
           f"&startDateDay={start_date.format('YYYY/MM/DD')}&maxIntradayDays=1&spanType=Intraday&"
           f"startDateIntraday={start_date.format('YYYY/MM/DD')}&startHourIntraday={start_date.hour}"
           f"&startMinuteIntraday={start_date.minute}&endDateIntraday={end_date.format('YYYY/MM/DD')}"
           f"&endHourIntraday={end_date.hour}&endMinuteIntraday={end_date.minute}&_adjustPlanHours=on"
           f"&_hideEmptyLineItems=on&employmentType=AllEmployees")

    with requests_retry_session() as req:
        resp = req.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL), verify=False,
                       allow_redirects=True, timeout=30)

    df_ppr = pd.read_csv(io.StringIO(resp.text))

    return df_ppr


def get_fc_tot(fc: str, start_date: pendulum, end_date: pendulum):
    url = (f"https://fclm-portal.amazon.com/reports/timeOnTask?reportFormat=CSV&warehouseId={fc}"
           f"&startDateDay={start_date.format('YYYY/MM/DD')}&maxIntradayDays=30&spanType=Intraday"
           f"&startDateIntraday={start_date.format('YYYY/MM/DD')}&startHourIntraday={start_date.hour}"
           f"&startMinuteIntraday={start_date.minute}&endDateIntraday={end_date.format('YYYY/MM/DD')}&"
           f"endHourIntraday={end_date.hour}&endMinuteIntraday={end_date.minute}")

    with requests_retry_session() as req:
        resp = req.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL), verify=False,
                       allow_redirects=True, timeout=30)

    df_tot = pd.read_csv(io.StringIO(resp.text))

    return df_tot


def get_function_rollup(fc: str, start_date: pendulum, end_date: pendulum, process_id: str):
    url = (f"https://fclm-portal.amazon.com/reports/functionRollup?reportFormat=CSV&warehouseId={fc}"
           f"&processId={process_id}&startDateDay={start_date.format('YYYY/MM/DD')}&maxIntradayDays=1"
           f"&spanType=Intraday&startDateIntraday={start_date.format('YYYY/MM/DD')}&startHourIntraday={start_date.hour}"
           f"&startMinuteIntraday={start_date.minute}&endDateIntraday={end_date.format('YYYY/MM/DD')}"
           f"&endHourIntraday={end_date.hour}&endMinuteIntraday={end_date.minute}")

    with requests_retry_session() as req:
        resp = req.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL), verify=False,
                       allow_redirects=True, timeout=30)

    df_fr = pd.read_csv(io.StringIO(resp.text))

    return df_fr


class HourlyTOT:
    def __init__(self, fc: str):
        self.fc = fc
        self.tz = timezone(self.fc)
        self.end_date = pendulum.today(self.tz).add(hours=pendulum.now(self.tz).hour)
        self.start_date = self.end_date.subtract(hours=1)
        self.df_roster = get_fc_roster(fc=self.fc)
        self.df_tot = self.flag_tot()
        self.df_tot_reminder = self.tot_history()
        self.df_ppr = get_ppr(self.fc, start_date=self.start_date, end_date=self.end_date)
        self.process_ids = {"Each-Receive - Total": 1003027, "LP Receive": 1003031,
                            "Each Transfer In - Total": 1002976, "Case Transfer In": 1003035,
                            "Pallet Transfer In - Total": 1003041, "Cubiscan": 1002971,
                            "Each Stow to Prime - Total": 1003016, "Prep - Total": 1003002}
        self.exclude_functions = ["Shipping Clerk"]
        self.df_zero = self.find_0_units()

        self.url = ("https://hooks.chime.aws/incomingwebhooks/258dfe1c-1585-4e9a-8359-4a35d2222257?token=NGNuZ21uVmN8MXxxM3RxdXA3eE85UXExemFFSXdsVnlmSHppU1VlTkxJREZZcU9KU3psb21V")

        if not self.df_zero.empty:
            self.send_report(self.df_zero,
                             f"Zero Units Processed {self.start_date.hour}-{self.end_date.hour}")
        if not self.df_tot.empty:
            self.send_report(self.df_tot, "Time off Task")

        if not self.df_tot_reminder.empty:
            self.send_report(self.df_tot_reminder, "Last 24 Hours TOT")

    def flag_tot(self):
        df = get_fc_tot(fc=self.fc, start_date=self.start_date, end_date=self.end_date)
        df = df[df["Percent Time on Task"] < 100].reset_index(drop=True)
        if df.empty:
            return df
        df = pd.merge(df, self.df_roster, how="left", on="Employee ID")
        df = df[["Employee Name_x", "User ID", "Manager", "Time On Task", "Total Time",
                "Percent Time on Task"]]
        df.rename(columns={"User ID": "Login", "Employee Name_x": "Employee Name"}, inplace=True)
        df = df.sort_values(["Manager", "Employee Name"], ascending=[1, 1]).reset_index(drop=True)
        df = df.dropna().reset_index(drop=True)
        if df.empty:
            return df
        df["Employee Name"] = df.apply(lambda x: self.add_waterfall(x, start=self.start_date, end=self.end_date),
                                       axis=1)
        return df

    def tot_history(self):
        df = get_fc_tot(fc=self.fc, start_date=self.start_date.subtract(days=1), end_date=self.start_date)
        df = df[df["Percent Time on Task"] < 100].reset_index(drop=True)
        if df.empty:
            return df
        df = pd.merge(df, self.df_roster, how="left", on="Employee ID")
        df = df[["Employee Name_x", "User ID", "Manager", "Time On Task", "Total Time",
                "Percent Time on Task"]]
        df.rename(columns={"Employee Name_x": "Employee Name", "User ID": "Login"}, inplace=True)
        df = df.sort_values(["Manager", "Employee Name"], ascending=[1, 1]).reset_index(drop=True)
        df = df.dropna().reset_index(drop=True)
        if df.empty:
            return df
        df["Employee Name"] = df.apply(lambda x: self.add_waterfall(x, start=self.start_date.subtract(days=1),
                                                                    end=self.start_date), axis=1)
        return df

    def find_0_units(self) -> pd.DataFrame:
        df_list = []

        for process_id in self.process_ids:
            units_processed = self.df_ppr.loc[self.df_ppr["LineItem Name"] == process_id, "Actual Hours"]
            if units_processed.empty:
                continue
            units_processed = units_processed.reset_index(drop=True)[0]
            if units_processed > 0:
                sleep(.5)
                df_list.append(self.get_0_processed(self.process_ids[process_id]))

        df = pd.concat(df_list).reset_index(drop=True)
        df = df[df["Units"] <= 0].reset_index(drop=True)
        df = df[df["Hours"] >= .16].reset_index(drop=True)
        df = pd.merge(df, self.df_roster, how="left", on="Employee ID")
        df = df[["Employee Name", "User ID", "Manager Name", "Hours", "Units", "Function Name"]]
        df.rename(columns={"User ID": "Login", "Manager Name": "Manager"}, inplace=True)
        df = df.sort_values(["Function Name", "Employee Name"], ascending=[1, 1]).reset_index(drop=True)
        df = df.dropna().reset_index(drop=True)
        if df.empty:
            return df
        df["Employee Name"] = df.apply(lambda x: self.add_waterfall(x, start=self.start_date, end=self.end_date),
                                       axis=1)
        return df

    def get_0_processed(self, process_id) -> pd.DataFrame:
        df = get_function_rollup(self.fc, start_date=self.start_date, end_date=self.end_date,
                                 process_id=process_id)
        df = df[df["Size"] == "Total"].reset_index(drop=True)
        df = df[df["Units"] <= 0].reset_index(drop=True)
        df = df[~df["Function Name"].isin(self.exclude_functions)].reset_index(drop=True)
        df = df[["Employee Id", "Name", "Paid Hours-Total(function,employee)", "Units", "Function Name"]]
        df.rename(columns={"Employee Id": "Employee ID", "Paid Hours-Total(function,employee)": "Hours"}, inplace=True)
        df = df.drop_duplicates(subset=["Employee ID", "Function Name"], keep="first").reset_index(drop=True)
        return df

    def add_waterfall(self, row, start: pendulum, end: pendulum) -> str:
        try:
            name = row["Employee Name"]
            login = row["Login"]
            url = (f"https://fclm-portal.amazon.com/employee/timeDetails?&employeeId={login.lower()}"
                   f"&warehouseId={self.fc}&startDateDay={start.format('YYYY/MM/DD')}&"
                   f"maxIntradayDays=1&spanType=Intraday&startDateIntraday={start.format('YYYY/MM/DD')}"
                   f"&startHourIntraday={start.hour}&startMinuteIntraday=0"
                   f"&endDateIntraday={end.format('YYYY/MM/DD')}&"
                   f"endHourIntraday={end.hour}&endMinuteIntraday=0")
            link = f"[{name.title()}]({url})"
            return link
        except AttributeError:
            return None

    def send_report(self, df: pd.DataFrame, title) -> None:
        size = 8
        report_dfs = [df.loc[i:i + size - 1, :] for i in range(0, len(df), size)]
        for df_report in report_dfs:
            url = f"https://fclm-portal.amazon.com/reports/timeOnTask?&warehouseId={self.fc}"
            mark_down = tabulate(df_report, tablefmt="github", headers="keys", showindex=False)
            message = (
                    f"/md ## [TOT]({url}) {self.fc} {title}: \n"
                    + "@Present"
                    + " \n"
                    + "--- \n"
                    + f"{mark_down}"
            )
            with requests_retry_session() as request:
                request.post(self.url, json={"Content": message})


def job_function():
    HourlyTOT("PSP1")


if __name__ == '__main__':
    sched = BlockingScheduler()
    sched.add_job(job_function, "cron", day_of_week="*", hour="*", minute="20", misfire_grace_time=600, coalesce=True)
    sched.start()
