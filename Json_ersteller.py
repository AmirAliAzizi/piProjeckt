import json, os

os.chdir("/home/pi/Desktop") #C:/Users/amir-ali.azizi/PycharmProjects/RD-TS/python-selenium-projekt

data ={"Pfade": ["http://vms075/teststatistik/#/",
                 "https://vms020user:Kiosk@intra.mybender.com/",
                 "https://vms020user:Kiosk@portal.mybender.com/sites/SoftwareQS/default.aspx",
                 "https://redmine.intra.bender/projects/software-qs-rd-ts-sil",
                 "https://vms020user:Kiosk@intra.mybender.com/de/Seiten/news-infos/KostBar.aspx",
                 "http://powerscout.bender.de"]}

with open ("Pfade_list_pi1.json","w") as json_file:
    json.dump(data, json_file, indent=2)
