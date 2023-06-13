import asyncio
import sys

from pyppeteer import connect, launch
import subprocess
import tkinter as tk
import re
import DolphinController


# изменить порядок сортировки


def remove_extra_spaces(text):
    t = re.sub(' +', ' ', text).strip()
    return re.sub(r"(\n\s*){2,}", "\n", t)


class Parser:
    def __init__(self, browser, token, pause, chrome_path):
        self.browser_name = browser
        self.chrome_path = chrome_path
        self.browser = None
        self.page = None
        self.current = 0
        self.controller = None
        self.profile_id = None
        self.token = token
        self.pause = pause

    async def open_page(self, url):
        root = tk.Tk()
        screen_width = 1520
        screen_height = 750
        root.destroy()
        if self.browser_name == "Chrome":
            if sys.platform == 'win32' or sys.platform == 'win64':
                subprocess.Popen([self.chrome_path, "--remote-debugging-port=9222"])
            elif sys.platform == "darwin":
                self.chrome_path = self.chrome_path.replace(" ", "\\ ")
                subprocess.Popen(
                    [self.chrome_path, "--remote-debugging-port=9222", "--no-first-run", "--no-default-browser-check",
                     "--user-data-dir=$(mktemp -d -t 'chrome-remote_data_dir')"])
            self.browser = await connect({"browserURL": "http://127.0.0.1:9222"})
        else:
            token = self.token
            self.controller = DolphinController.Controller(token)
            profile_id = self.controller.create_profile_and_send_id()
            automation_json = self.controller.start_automation()
            dolphin_port = automation_json["automation"]["port"]
            dolphin_ws_endpoint = automation_json["automation"]["wsEndpoint"]
            self.browser = await connect({"browserWSEndpoint": f'ws://127.0.0.1:{dolphin_port}{dolphin_ws_endpoint}'})
        self.page = await self.browser.newPage()
        await self.page.setViewport({"width": screen_width, "height": screen_height})
        await self.page.goto(url)
        self.page.setDefaultNavigationTimeout(120000)

    async def parse(self, fr, to, act_type, court):
        await asyncio.sleep(self.pause)

        if court != "":
            court_input = await self.page.waitForSelector('#caseCourt')
            await court_input.click()
            await asyncio.sleep(0.5)
            await self.page.keyboard.type(court)
            await asyncio.sleep(0.5)
            await self.page.mouse.click(0, 300)
            await asyncio.sleep(0.5)

        await self.page.waitForSelector('#sug-dates')
        inp = (await self.page.querySelectorAll("#sug-dates input"))[0]
        await inp.click()
        await asyncio.sleep(1)
        await self.page.keyboard.type(fr)
        await asyncio.sleep(0.5)
        inp_to = (await self.page.querySelectorAll("#sug-dates input"))[1]
        await inp_to.click()
        await asyncio.sleep(0.5)
        await self.page.keyboard.type(to)
        await asyncio.sleep(0.5)
        await self.page.keyboard.press("Enter")
        await self.page.keyboard.up("Enter")
        keyboard = self.page.keyboard
        await asyncio.sleep(1.5)
        if act_type == 1:
            await self.page.waitForSelector(".administrative")
            btn = (await self.page.querySelectorAll(".administrative"))[0]
            await btn.click()
        elif act_type == 2:
            await self.page.waitForSelector(".civil")
            btn = (await self.page.querySelectorAll(".civil"))[0]
            await btn.click()
        elif act_type == 3:
            await self.page.waitForSelector(".bankruptcy")
            btn = (await self.page.querySelectorAll(".bankruptcy"))[0]
            await btn.click()
        await asyncio.sleep(1.5)

        await self.page.waitForXPath('//div[@class="b-case-loading" and @style="display: none;"]')

        await self.page.waitForXPath('//*[@id="pages"]//li[not(@class)]', {'timeout': 180000})
        pages = await self.page.xpath('//*[@id="pages"]//li[not(@class)]')
        num_of_pages = await (await pages[-1].querySelector('a')).getProperty('text')
        num_of_pages = await num_of_pages.jsonValue()
        num_of_pages = int(num_of_pages)

        main_string = """"""
        main_list = []
        prev_num_case = ""
        while self.current < num_of_pages:
            await self.page.waitForXPath('//div[@class="b-case-loading" and @style="display: none;"]')
            # while remove_extra_spaces(await (
            #         await (await self.page.querySelector(".num_case")).getProperty(
            #             "text")).jsonValue()) == prev_num_case:
            #     continue
            # prev_num_case = remove_extra_spaces(
            #     await (await (await self.page.querySelector(".num_case")).getProperty("text")).jsonValue())
            await self.page.waitForSelector('.num_case')
            lines = await self.page.querySelectorAll('#b-cases tr')
            for line in lines:
                try:
                    current_dict = {}

                    case = await line.querySelector('.num_case')

                    num_case_text = await (await case.getProperty("text")).jsonValue()
                    num_case_text = remove_extra_spaces(num_case_text)

                    num_case_link = await (await case.getProperty('href')).jsonValue()

                    case_date = await line.querySelector('.num span')
                    case_date = await (await case_date.getProperty('innerText')).jsonValue()

                    case_judge = await line.querySelector('.court .judge')
                    case_judge = await (await case_judge.getProperty('innerText')).jsonValue()

                    case_court = await line.querySelector('.court .b-container')
                    case_court = (await case_court.querySelectorAll("*"))[-1]
                    case_court = await (await case_court.getProperty('innerText')).jsonValue()

                    current_dict["case_date"] = case_date
                    current_dict["case_type"] = act_type
                    current_dict["case_num"] = num_case_text
                    current_dict["case_link"] = num_case_link
                    current_dict["case_judge"] = case_judge
                    current_dict["case_court"] = case_court
                    # case_str = f"Дело№: {num_case_text}; Ссылка: {num_case_link}\n\n"

                    plaintiffs = await line.querySelector('.plaintiff')
                    plaintiffs = await plaintiffs.querySelectorAll('.js-rolloverHtml')
                    respondents = await line.querySelector(".respondent")
                    respondents = await respondents.querySelectorAll(".js-rolloverHtml")

                    plaintiffs_list = []
                    for plaintiff in plaintiffs:
                        plaintiff_name = await plaintiff.querySelector('strong')
                        plaintiff_name = await (await plaintiff_name.getProperty("innerText")).jsonValue()
                        plaintiff_text = await (await plaintiff.getProperty("innerText")).jsonValue()
                        plaintiff_text = remove_extra_spaces(plaintiff_text)
                        plaintiff_text = plaintiff_text.replace(plaintiff_name, "")
                        plaintiff_text = plaintiff_text[1:]

                        # print(plaintiff_text)
                        current_plaintiff_dict = {"plaintiff_name": plaintiff_name, "plaintiff_address": plaintiff_text}
                        plaintiffs_list.append(current_plaintiff_dict)
                        # case_str += f"Истец: {plaintiff_text}\n"

                    # if len(plaintiffs) != 0:
                    #     case_str += "\n"
                    # else:
                    #     case_str += "Истец: нет информации\n\n"

                    respondent_list = []
                    for respondent in respondents:
                        respondent_name = await respondent.querySelector('strong')
                        respondent_name = await (await respondent_name.getProperty("innerText")).jsonValue()
                        respondent_text = await (await respondent.getProperty("innerText")).jsonValue()
                        respondent_text = remove_extra_spaces(respondent_text)
                        respondent_text = respondent_text.replace(respondent_name, "")
                        respondent_text = respondent_text[1:]

                        # print(respondent_text)
                        current_respondent_dict = {"respondent_name": respondent_name,
                                                   "respondent_address": respondent_text}
                        respondent_list.append(current_respondent_dict)
                        # case_str += f"Ответчик: {respondent_text}\n"

                    # if len(respondents) != 0:
                    #     case_str += "\n\n"
                    # else:
                    #     case_str += "Ответчик: нет информации\n\n\n"

                    # main_string += case_str
                    current_dict["plaintiff"] = plaintiffs_list
                    current_dict["respondent"] = respondent_list
                    main_list.append(current_dict)
                except Exception as e:
                    print(e)

            self.current = self.current + 1
            await keyboard.down("Control")
            await keyboard.press("ArrowRight")
            await keyboard.up("ArrowRight")
            await keyboard.up("Control")

        try:
            if self.browser_name == "Dolphin":
                self.controller.stop_profile()
                self.controller.delete_browser_profile()
            await self.page.close()
            await self.browser.close()
        except Exception as e:
            print(e)
        # return main_string
        return main_list

    async def start_parse(self, url, fr, to, act_type, court):
        await self.open_page(url)
        main_string = await self.parse(fr, to, act_type, court)
        return main_string
