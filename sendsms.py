from ipaddress import IPv4Address  # for your IP address
import csv 
import time
import random
from pyairmore.request import AirmoreSession  # to create an AirmoreSession
from pyairmore.services.messaging import MessagingService  # to send messages


def Connect():
    ip = IPv4Address("192.168.1.68")  # let's create an IP address object
    # now create a session
    session = AirmoreSession(ip)
    # if your port is not 2333
    # session = AirmoreSession(ip, 2334)  # assuming it is 2334
    session.is_server_running  # True if Airmore is running
    was_accepted = session.request_authorization()
    print(was_accepted)  # True if accepted
    return session


def ReadCSV(FileName):
    data_list = []
    with open(f"{FileName}", 'r') as file:
        csvreader = csv.reader(file)
        for row in csvreader:
            data_list.append(row)
    return data_list

def main():
    file = input("Enter the file path .csv: ")
    datas =  ReadCSV(file)
    session = Connect() 
    body1 = "बिधयालय लाई आवश्यक पर्ने कम्प्युटर, ल्यपटप, इन्टेरयाकटिब प्यानल बोर्ड, स्मार्ट टिभी, साइन्स ल्याब सामाग्री हरु एकै छातामुनी सर्बसुलभ रुपमा उपलब्ध छ ।"
    body2 ="सम्पर्क : मेगाटेक, मेनरोड, \n बिराटनगर \n ९८२०७५६२८०, ९८२०७५६२८२ \n हजुरकोको खुशी हाम्रो चाहना \n Facebook page: http://pili.app "
    for data in datas:
        service = MessagingService(session)
        service.send_message(data[2], body1)
        time.sleep(random.randrange(1,4))
        service.send_message(data[2], body2)


if __name__=="__main__":
    main()