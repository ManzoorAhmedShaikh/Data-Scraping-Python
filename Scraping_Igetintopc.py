from bs4 import BeautifulSoup
import requests
import gspread
from gspread_wrappers import GspreadWrapper

class Igetintopc_Data:

    credential_file = "" #Add the path of your credential file for Gspread connectivity
    Sheet_name = ""      #Add the sheet name that is in your Google sheets
    def extract(self,number_of_pages):
        """
        It will extract the Igetintopc.com app data to the spreadsheet
        :param number_of_pages: A positive number greater than 1 till that the data would scrap
        """

        #Initialze the Google sheet and Gspread wrapper
        self.ws = gspread.service_account(filename=self.credential_file).open(self.Sheet_name).get_worksheet(0)
        self.gsw = GspreadWrapper()

        #Styling of your sheet
        Head_Style = self.gsw.CreateFormattingStyle(bg=[1,1,0.25],fsize=13,border=True)
        Data_Style = self.gsw.CreateFormattingStyle(fsize=9,border=True)

        self.gsw.ApplyingMultiFormatting(self.ws,[("A1:E1",Head_Style),("A2:E",Data_Style)])
        self.gsw.UpdateCells(self.ws,"A1:E1",[["ID","App Name","Url","Image Url","Description"]])

        ids = 0
        final_list = []

        for x in range(1, number_of_pages+1):
            if x == 1:
                urls = f'https://igetintopc.com'
            else:
                urls = f'https://igetintopc.com/page/{x}/'

            Doc = requests.get(urls)
            Con = Doc.content
            soup_obj = BeautifulSoup(Con, "lxml")
            np = int(soup_obj.find('a',{'class':"last"}).text) #Total number of pages that are available in Igetintopc

            if number_of_pages > np:
                print(f"Sorry you have exceeded the total number of pages that are {np}. Try any other number less than that!")
                break

            elements = soup_obj.find('div', {"class": "posts clear-block"})

            for y in elements:
                try:
                    if len(y) >= 3:
                        ids += 1
                        Name = y.find_next('h2', {'class': 'title'}).find('a').text
                        Link = y.find_next('h2', {'class': 'title'}).find('a')['href']
                        Img = y.find_next('img')['src']
                        Descp = y.find_next('div', {'class': 'post-content clear-block'}).text

                        final_list.append([ids,Name,Link,Img,Descp])
                    else:
                        pass
                except:
                    pass

        #Adding data to the sheets
        self.gsw.UpdateCells(self.ws,"A2:E",final_list)

    def clear_sheet(self):
        """
        It will clear all the sheet data
        :return:
        """
        rngs = self.ws.range("A2:E")
        for x in rngs:
            x.value = ""

        self.ws.update_cells(rngs)