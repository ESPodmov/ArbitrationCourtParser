import json
import random
import requests

platforms_array = ["macos", "windows", "linux"]
vendors_array = ["Google Inc. (Intel)", "Google Inc. (NVIDIA)", "Google Inc. (AMD)", "Google Inc.",
                 "Google Inc. (Microsoft)", "Google Inc. (Unknown)", "Intel Inc.",
                 "Google Inc. (NVIDIA Corporation)"]
renderers_array = {
    "Google Inc. (Intel)": [
        "ANGLE (Intel, Intel(R) HD Graphics Family Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 600 Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.100.7870)",
        "ANGLE (Intel, Intel(R) HD Graphics 620 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 620 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 600 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics 520 Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8681)",
        "ANGLE (Intel, Intel(R) Iris(TM) Graphics 5100 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics 4000 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 620 Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8681)",
        "ANGLE (Intel, Intel(R) HD Graphics 4600 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 600 Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8681)",
        "ANGLE (Intel, Intel(R) HD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11-23.20.16.4974)",
        "ANGLE (Intel, Intel(R) HD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-10.18.10.4276)",
        "ANGLE (Intel, Intel(R) UHD Graphics 620 Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.100.7925)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.100.7985)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8476)",
        "ANGLE (Intel, Intel(R) HD Graphics 530 Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.100.8142)",
        "ANGLE (Intel, Intel(R) HD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) Iris(R) Xe Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8935)",
        "ANGLE (Intel, Intel(R) HD Graphics 610 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics Family Direct3D11 vs_5_0 ps_5_0, D3D11-9.18.13.2739)",
        "ANGLE (Intel, Intel(R) UHD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.101.1122)",
        "ANGLE (Intel, Intel(R) UHD Graphics 620 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.100.9864)",
        "ANGLE (Intel, Intel(R) HD Graphics 520 Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8682)",
        "ANGLE (Intel, Intel(R) UHD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11-25.20.100.6446)",
        "ANGLE (Intel, Intel(R) HD Graphics 5500 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8681)",
        "ANGLE (Intel, Intel(R) HD Graphics 4600 Direct3D11 vs_5_0 ps_5_0, D3D11-10.18.14.5067)",
        "ANGLE (Intel, Intel(R) Iris(R) Plus Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.9621)",
        "ANGLE (Intel, Intel(R) HD Graphics 530 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-10.18.10.4425)",
        "ANGLE (Intel, Intel(R) HD Graphics 520 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 730 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) HD Graphics 4400 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics 605 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.101.1191)",
        "ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.100.8783)",
        "ANGLE (Intel, Intel(R) HD Graphics 530 Direct3D11 vs_5_0 ps_5_0, D3D11-20.19.15.4404)",
        "ANGLE (Intel, Intel(R) HD Graphics 530 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.101.1191)",
        "ANGLE (Intel, Intel(R) HD Graphics 4600 Direct3D11 vs_5_0 ps_5_0, D3D11-9.18.10.3215)"
    ],
    "Google Inc. (NVIDIA)": [
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1060 6GB Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7247)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1050 Ti Direct3D11 vs_5_0 ps_5_0, D3D11-27.21.14.6172)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 3060 Laptop GPU Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.9729)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 2070 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7196)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 3070 Ti Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.15.1179)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 965M Direct3D11 vs_5_0 ps_5_0, D3D11-26.21.14.3086)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1070 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7196)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1660 Ti Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.9733)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX TITAN Black Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7298)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 960 Direct3D11 vs_5_0 ps_5_0, D3D11-27.21.14.6109)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1060 6GB Direct3D11 vs_5_0 ps_5_0, D3D11-27.21.14.6663)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1080 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7111)",
        "ANGLE (NVIDIA, NVIDIA GeForce 315M    Direct3D11 vs_4_1 ps_4_1, D3D11-9.18.13.4192)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 2070 SUPER Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7247)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1060 Direct3D11 vs_5_0 ps_5_0, D3D11-26.21.14.3190)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 970 Direct3D11 vs_5_0 ps_5_0, D3D11-25.21.14.1681)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1050 Ti Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7141)",
        "ANGLE (NVIDIA, NVIDIA GeForce 940MX Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (NVIDIA, NVIDIA GeForce GT 640 Direct3D11 vs_5_0 ps_5_0, D3D11-27.21.14.5730)",
        "ANGLE (NVIDIA, NVIDIA Quadro K4100M Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 2060 Direct3D11 vs_5_0 ps_5_0, D3D11-26.21.14.4223)",
        "ANGLE (NVIDIA, NVIDIA GeForce RTX 2060 SUPER Direct3D11 vs_5_0 ps_5_0, D3D11-27.21.14.5751)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1080 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.9676)",
        "ANGLE (NVIDIA, NVIDIA Quadro K620 Direct3D11 vs_5_0 ps_5_0, D3D11-10.18.13.5330)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1050 Ti Direct3D11 vs_5_0 ps_5_0, D3D11-26.21.14.3107)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1650 SUPER Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.15.1123)",
        "ANGLE (NVIDIA, NVIDIA GeForce GT 640 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7212)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 1650 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7280)",
        "ANGLE (NVIDIA, NVIDIA GeForce GTX 660 Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14.7141)",
        "ANGLE (NVIDIA, NVIDIA Quadro P2000 Direct3D11 vs_5_0 ps_5_0, D3D11-26.21.14.3086)",
        "ANGLE (NVIDIA, NVIDIA Quadro 3000M Direct3D11 vs_5_0 ps_5_0, D3D11)"
    ],
    "Google Inc. (AMD)": [
        "ANGLE (AMD, AMD Radeon(TM) Vega 10 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon(TM) Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.13044.0)",
        "ANGLE (AMD, AMD Radeon RX 6800 XT Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon(TM) Vega 8 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-23.20.840.0)",
        "ANGLE (AMD, AMD Radeon(TM) Vega 3 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.14544.6)",
        "ANGLE (AMD, AMD Radeon(TM) RX Vega 10 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.11030.22001)",
        "ANGLE (AMD, Radeon RX 580 Series Direct3D11 vs_5_0 ps_5_0, D3D11-30.0.14029.1011)",
        "ANGLE (AMD, AMD Radeon(TM) R4 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-21.19.144.1281)",
        "ANGLE (AMD, AMD Radeon HD 8400 / R3 Series Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.1034.6)",
        "ANGLE (AMD, AMD Radeon(TM) Vega 8 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.12033.1007)",
        "ANGLE (AMD, AMD Radeon(TM) Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.11028.5001)",
        "ANGLE (AMD, AMD Radeon(TM) Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.11020.10001)",
        "ANGLE (AMD, AMD Radeon (TM) R5 M320 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon(TM) Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.11044.7)",
        "ANGLE (AMD, AMD Radeon R9 200 Series Direct3D11 vs_5_0 ps_5_0, D3D11-22.19.171.1024)",
        "ANGLE (AMD, AMD Radeon(TM) Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.21034.37)",
        "ANGLE (AMD, AMD Radeon(TM) Vega 3 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.15030.5001)",
        "ANGLE (AMD, AMD Radeon R7 200 Series Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.15029.20013)",
        "ANGLE (AMD, AMD Radeon R9 200 Series Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.21003.8013)",
        "ANGLE (AMD, AMD Radeon(TM) Vega 8 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.21034.37)",
        "ANGLE (AMD, AMD Radeon(TM) R3 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-26.20.12036.8001)",
        "ANGLE (AMD, Radeon (TM) Pro WX 4150 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon R7 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon (TM) R5 M330 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon (TM) R7 360 Series Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.21003.8013)",
        "ANGLE (AMD, AMD Radeon HD 5670 Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (AMD, AMD Radeon(TM) R5 Graphics Direct3D11 vs_5_0 ps_5_0, D3D11-23.20.808.3840)",
        "ANGLE (AMD, AMD Radeon (TM) R7 370 Series Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.21026.2006)"
    ],
    "Google Inc.": [
        "ANGLE (Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (Intel(R) UHD Graphics 620 Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (NVIDIA GeForce GT 440  Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (AMD Radeon(TM) Vega 8 Graphics Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (Intel(R) HD Graphics Direct3D11 vs_4_1 ps_4_1)",
        "ANGLE (AMD Radeon(TM) R7 Graphics Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (Intel(R) UHD Graphics 600 Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (Intel(R) HD Graphics 4400 Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (Intel(R) HD Graphics 4000 Direct3D11 vs_5_0 ps_5_0)",
        "ANGLE (NVIDIA GeForce GTX 750 Direct3D11 vs_5_0 ps_5_0)"
    ],
    "Google Inc. (Microsoft)": [
        "ANGLE (Microsoft, Microsoft Basic Render Driver Direct3D11 vs_5_0 ps_5_0, D3D11-10.0.19041.546)",
        "ANGLE (Microsoft, Microsoft Basic Render Driver Direct3D11 vs_5_0 ps_5_0, D3D11-6.3.9600.16505)"
    ],
    "Google Inc. (Unknown)": [
        "ANGLE (Unknown, Qualcomm(R) Adreno(TM) 630 GPU Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Unknown, Qualcomm(R) Adreno(TM) 630 GPU Direct3D11 vs_5_0 ps_5_0, D3D11-25.18.10500.0)",
        "ANGLE (Unknown, Parallels Display Adapter (WDDM) Direct3D11 vs_5_0 ps_5_0, D3D11)",
        "ANGLE (Unknown, Qualcomm(R) Adreno(TM) 680 GPU Direct3D11 vs_5_0 ps_5_0, D3D11-27.20.1640.0)",
        "ANGLE (Unknown, VMware SVGA 3D Direct3D11 vs_4_1 ps_4_1, D3D11-8.16.1.24)"
    ],
    "Intel Inc.": [
        "Intel Iris OpenGL Engine"
    ],
    "Google Inc. (NVIDIA Corporation)": [
        "ANGLE (NVIDIA Corporation, GeForce RTX 3070/PCIe/SSE2, OpenGL 4.5.0 NVIDIA 461.40)"
    ]
}
cpu = [2, 4, 6, 8, 12, 16]
memory = [2, 4, 8]


def authentication(username: str, password: str):
    options = {
        "url": "https://anty-api.com/auth/login",
        "data": {
            "username": username,
            "password": password
        }
    }
    response = requests.post(options["url"], data=options["data"])

    response_json = response.json()
    # print(str(response_json))
    if "token" in response_json:
        return response_json["token"]
    else:
        raise Exception("Authentication error, check your Internet connection and config data")


class Controller:
    def __init__(self, token):
        auth_token = token
        self.headers = {
            "Authorization": f"Bearer {auth_token}",
            "Content-Type": "application/json"
        }
        self.profile_id = None

    def get_new_user_agent(self, platform: str):
        session = requests.session()
        session.headers = self.headers
        version_number = random.randint(101, 107)
        options = {
            "url": f"https://anty-api.com/fingerprints/useragent?browser_type=anty&browser_version=107&" +
                   f"platform={platform}"
        }
        response = session.get(options["url"]).json()
        session.close()
        # print(response)
        if "data" in response:
            return response["data"]
        else:
            raise Exception("Can't get new user agent, something went wrong")

    def create_profile_and_send_id(self):
        session = requests.session()
        session.headers = self.headers
        null = None
        # set current parameters for browser profile
        current_vendor = random.choice(vendors_array)
        current_renderer = random.choice(renderers_array[current_vendor])
        current_platform = random.choice(platforms_array)
        current_user_agent = self.get_new_user_agent(platform=current_platform)
        # set options for browser profile
        options = {
            "url": "https://anty-api.com/browser_profiles",
            "data": {
                "platform": current_platform,
                "browserType": "anty",
                "useragent": {
                    "mode": "manual",
                    "value": current_user_agent
                },
                "webrtc": {
                    "mode": "altered",
                    "ipAddress": null
                },
                "canvas": {
                    "mode": "real"
                },
                "webgl": {
                    "mode": "real"
                },
                "webglInfo": {
                    "mode": "manual",
                    "vendor": current_vendor,
                    "renderer": current_renderer
                },
                "geolocation": {
                    "mode": "real",
                    "latitude": null,
                    "longitude": null
                },
                "cpu": {
                    "mode": "manual",
                    "value": random.choice(cpu)
                },
                "memory": {
                    "mode": "manual",
                    "value": random.choice(memory)
                },
                "timezone": {
                    "mode": "auto",
                    "value": null
                },
                "locale": {
                    "mode": "auto",
                    "value": null
                },
                "name": "Current Bot",
                "tags": [
                    ""
                ],
                "mainWebsite": "",
                "notes": {
                    "content": null,
                    "color": "blue",
                    "style": "text",
                    "icon": null
                },
                "proxy": null,
                "statusId": 0,
                "doNotTrack": False
            }
        }
        response = session.post(options["url"], data=json.dumps(options["data"]))
        # print(response.json())
        profile_id = self.get_last_browser_profile()
        session.close()
        self.profile_id = profile_id['id']
        return profile_id['id']

    def get_last_browser_profile(self):
        session = requests.session()
        session.headers = self.headers
        options = {
            "url": "https://anty-api.com/browser_profiles"
        }
        response = session.get(options["url"]).json()
        session.close()
        if "data" in response:
            return response["data"][0]
        else:
            raise Exception("Can't get list of browser profiles")

    def delete_browser_profile(self, profile_id: str):
        session = requests.session()
        session.headers = self.headers
        options = {
            'url': "https://anty-api.com/browser_profiles/"
        }
        response = session.delete(options["url"] + str(profile_id))
        session.close()
        # print(response.json())

    def start_automation(self):
        profile_id = self.profile_id
        session = requests.session()
        session.headers = self.headers
        options = {
            "url": f"http://localhost:3001/v1.0/browser_profiles/{profile_id}/start?automation=1"
        }
        response = session.get(options['url']).json()
        session.close()
        return response

    def stop_profile(self):
        profile_id = self.profile_id
        session = requests.session()
        session.headers = self.headers
        options = {
            "url": f"http://localhost:3001/v1.0/browser_profiles/{profile_id}/stop"
        }
        response = session.get(options["url"]).json()
        session.close()
        return response
