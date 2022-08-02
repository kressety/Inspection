from sqlite3 import connect

from bs4 import BeautifulSoup
from requests import get
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects, HTTPError
from re import findall

RequestHeader = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/103.0.5060.134 Safari/537.36 Edg/103.0.1264.77 '
}


def GetFromUrl(Url):
    try:
        Request = get(Url, headers=RequestHeader)
        if Request.status_code == 200:
            HTMLData = BeautifulSoup(Request.text, features='lxml')
            return HTMLData
        else:
            return '{}: {}'.format(Request.status_code, Request.reason)
    except ConnectionError:
        return '网络连接错误'
    except Timeout:
        return '请求超时'
    except TooManyRedirects:
        return '重定向次数过多'
    except HTTPError:
        return '网页请求错误'


def Gaokao_Scheme():
    Name = '招生计划专刊'
    Url = 'https://www.jseea.cn/webfile/zkzl/gaokao/'
    Request = GetFromUrl(Url)
    if type(Request) != str:
        BufferDB = connect('buffer')
        ToastCollection = []
        for Item in Request.find('ul', {'class': 'content-list-ul news-list'}).find_all('a'):
            WriteKey = Item.text.split(' ')
            WriteHref = 'http:{}'.format(Item.attrs['href'])
            BufferTable = BufferDB.execute('select Name from Gaokao_Scheme_Buffer')
            for BufferRow in BufferTable:
                if WriteKey[0] == BufferRow[0]:
                    break
            else:
                BufferDB.execute("insert into Gaokao_Scheme_Buffer (Name, Date, Href) \
                VALUES ('{}', '{}', '{}')".format(WriteKey[0], WriteKey[1], WriteHref))
                BufferDB.commit()
                ToastCollection.append(
                    ['Inspection - {}'.format(Name), '已发现更新: {}'.format(WriteKey[0]), WriteHref])
        BufferDB.close()
        return ToastCollection
    else:
        return Request


def Gaokao_Line():
    Name = '投档线'
    Url = 'https://www.jseea.cn/search/?wd=%E6%8A%95%E6%A1%A3%E7%BA%BF&orderField=publishDate'
    Request = GetFromUrl(Url)
    if type(Request) != str:
        BufferDB = connect('buffer')
        ToastCollection = []
        for Item in Request.find_all('div', {'class': 'search-result-item'}):
            WriteName = Item.h3.find_all('a')[1].text.strip()
            WriteHref = 'http:{}'.format(Item.h3.find_all('a')[1].attrs['href'])
            BufferTable = BufferDB.execute('select Name from Gaokao_Line_Buffer')
            for BufferRow in BufferTable:
                if WriteName == BufferRow[0]:
                    break
            else:
                BufferDB.execute("insert into Gaokao_Line_Buffer (Name, Href) \
                VALUES ('{}', '{}')".format(WriteName, WriteHref))
                BufferDB.commit()
                ToastCollection.append(['Inspection - {}'.format(Name), '已发现更新: {}'.format(WriteName), WriteHref])
        BufferDB.close()
        return ToastCollection
    else:
        return Request


def CSP_Notification():
    Name = 'CSP通知'
    BaseUrl = 'https://www.cspro.org'
    JumpRequest = GetFromUrl(BaseUrl + '/cms/show.action?code=jumpchanneltemplate')
    JumpParams = [Item.strip('\"') for Item in findall('\".*\"', JumpRequest.html.head.script.text)]
    JumpHeader = JumpParams[1] + JumpParams[0]
    Request = GetFromUrl(BaseUrl + JumpHeader)
    if type(Request) != str:
        BufferDB = connect('buffer')
        ToastCollection = []
        for Item in Request.find_all('span', {'class': 'l_newsmsg_title'}):
            WriteName = Item.text.strip()
            WriteHref = BaseUrl + Item.a.attrs['href']
            BufferTable = BufferDB.execute('select Name from CSP_Notification_Buffer')
            for BufferRow in BufferTable:
                if WriteName == BufferRow[0]:
                    break
            else:
                BufferDB.execute("insert into CSP_Notification_Buffer (Name, Href) \
                    VALUES ('{}', '{}')".format(WriteName, WriteHref))
                BufferDB.commit()
                ToastCollection.append(['Inspection - {}'.format(Name), '已发现更新: {}'.format(WriteName), WriteHref])
        BufferDB.close()
        return ToastCollection
    else:
        return Request


TaskList = [
    Gaokao_Scheme,
    Gaokao_Line,
    CSP_Notification
]
