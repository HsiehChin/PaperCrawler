import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
from openpyxl import load_workbook

# input file and output file accpept excel file
input_excel_name = "your excel file.xlsx"
output_excel_name = "your output excel file.xlsx"
sheet_name = "your sheet name"

paper_root_href = "https://ndltd.ncl.edu.tw" # 國圖搜尋主網址

options = Options()
options.add_argument("--incognito")
options.add_argument("--disable-notifications")
options.add_argument('--proxy-server=socks5://localhost:9050')

# Read the name of authors and title of paper
def read_excel(excel_name):
    print("Reading the excel:", excel_name, ".....\n")
    # return excel name of authors and title of paper
    df = pd.read_excel(excel_name)
    return df

def write_excel(df, output_excel_name, sheet_name):
    print("Writing the excel:", output_excel_name, ".....\n")
    writer = pd.ExcelWriter(output_excel_name, engine='openpyxl')
    df.to_excel(writer,index=False,header=True, sheet_name=sheet_name)
    writer.close()
    print("Successfuly write the excel:", output_excel_name, ".....\n")

################################################
# get_paper_contents
#   To retrieve the specific content on the result website
#   note: need to access more detailed path, or the content sometime will loss <br> format.
################################################

def get_paper_contents(chrome, paper_herf):
    print("Retrieve the paper contents....")
    paper_titles = ["研究生(外文)", "論文名稱(外文)", "中文關鍵詞", "外文關鍵詞", "摘要", "外文摘要", "目次", "參考文獻"]
    paper_contents = [["英作者",""], ["英書名",""], ["關鍵字",""], ["英關鍵字",""], ["摘要",""], ["英摘要",""], ["目次",""], ["參考文獻",""], ["名稱可能有誤", ""]]
    paper_contents_exists = []
    # excel index : 3 5 6 7 8 9 10 11
    try:
        chrome.get(paper_herf)
        sleep(0.2)
        titles = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]')
        sleep(0.2)
        contents = chrome.find_elements(By.XPATH, '//div[@id="gs32_levelrecord"]/div/div')
        sleep(0.2)

        if titles:
            print("Find the titles!")
            # 確認有沒有存在相關文件
            for i in paper_titles:
                if i in titles.text:
                    paper_contents_exists.append(True)       
                else:
                    paper_contents_exists.append(False)       
            print(paper_contents_exists)  


        if contents:
            print("Find the contents! Total: ", len(contents))
            # 論文基本資料，一定會有這一頁，所以不用做確認
            basic_info = contents[0].text.split('\n')
            
            # Access "研究生(外文)", "論文名稱(外文)", "中文關鍵詞", "外文關鍵詞"
            for ti in range(4): 
                for info in basic_info:
                    if paper_titles[ti] in info:
                        paper_contents[ti][1] = info.split(":")[1].lstrip()

            # 如果有摘要、外文摘要、目次、參考文獻
            access_index = 1
            if paper_contents_exists[4]: # 摘要
                li = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//ul[@class="yui-nav"]/li//a[@title="摘要"]')
                li.click()
                sleep(0.2)
                li_content = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//div[@class="yui-content"]//div[@style="display: block;"]//td[@class="stdncl2"]')
                paper_contents[4][1] = li_content.get_attribute("innerText").lstrip()
                access_index+=1
            
            if paper_contents_exists[5]: # 外文摘要
                li = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//ul[@class="yui-nav"]/li//a[@title="外文摘要"]')
                li.click()
                sleep(0.2)
                li_content = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//div[@class="yui-content"]//div[@style="display: block;"]//td[@class="stdncl2"]')
                paper_contents[5][1] = li_content.get_attribute("innerText").lstrip()
                access_index+=1
            
            if paper_contents_exists[6]: # 目次
                li = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//ul[@class="yui-nav"]/li//a[@title="目次"]')
                li.click()
                sleep(0.2)
                li_content = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//div[@class="yui-content"]//div[@style="display: block;"]//td[@class="stdncl2"]')
                paper_contents[6][1] = li_content.get_attribute("innerText").lstrip()
                access_index+=1
            
            if paper_contents_exists[7]: # 參考文獻
                li = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//ul[@class="yui-nav"]/li//a[@title="參考文獻"]')
                li.click()
                sleep(0.2)
                li_content = chrome.find_element(By.XPATH, '//div[@id="gs32_levelrecord"]//div[@class="yui-content"]//div[@style="display: block;"]//td[@class="stdncl2"]')
                paper_contents[7][1] = li_content.get_attribute("innerText").lstrip()


            # print(paper_contents[:])
    except:
        print("Something Wrong!")
    return paper_contents

################################################
# paper_check
#   This is a rough comparision algorithm to check the paper title and school on search result website.
#   1. check school name
#   2. Compare the paper first word and last word 
#   3. if 2 is not the same, then calculate the finding proportion is large than 0.5
#      if true, than the paper title is what we find.
################################################
def paper_check(search_list, i, author_status, author_name, paper_name, mode):
    result = False
    p_c = search_list[i*7+1].get_attribute('innerText')
    a_n = search_list[i*7+2].get_attribute('innerText')
    p_t = search_list[i*7].get_attribute('innerText') # find title
    if author_status in p_c and ("國立臺灣科技大學" in p_c or "國立台灣工業技術學院" in p_c):
        print(i, p_c, a_n, p_t)
        # Check the title correct
        total_paper_length = len(paper_name)
        correct_word = 0
        if (p_t == paper_name) or (p_t[0] == paper_name[0] and p_t[len(p_t)-1] == paper_name[len(paper_name)-1]):                        
            result = True
        else:
            for c in paper_name:
                if c in p_t:
                    correct_word +=1
            if total_paper_length > 0 and correct_word/total_paper_length >= 0.5:
                result = True                 

    return result

################################################
# search_by_paper_name
#   search the paper by paper name, 
#   1. Enter the paper name.
#   2. Check the search result content including the paper name
#       if the count of result large than 100, use the school filter to find specific result.
#       else find and enter the paper content website.
################################################

def search_by_paper_name(chrome, author_name, paper_name, author_status):
    print("Search the paper by paper title....")
    chrome.get(paper_root_href)
    try:
        qs0 = chrome.find_element("id", "ysearchinput0")
        qs0.send_keys(paper_name)
        qs0.submit()
        sleep(0.1)

        search = chrome.find_element(By.XPATH, '//td[@class="tdfmt1-content"]')
        sub_button = chrome.find_element(By.XPATH, '//div[@id="researchdivid"]/input')
        if search:
            search_list = chrome.find_elements(By.XPATH, '//td[@class="tdfmt1-content"]/div/div/table/tbody//td[@headers="simplefmt2td"]')
            search_account = int(len(search_list)/7)
            if search_account >= 30:
                research = chrome.find_element("id", "research")
                research.send_keys("國立臺灣科技大學")
                sub_button.click()
                sleep(0.2)

        search = chrome.find_element(By.XPATH, '//td[@class="tdfmt1-content"]')
        if search:            
            search_list = chrome.find_elements(By.XPATH, '//td[@class="tdfmt1-content"]/div/div/table/tbody//td[@headers="simplefmt2td"]')
            search_account = int(len(search_list)/7)
            access_index = -1
            print("Search result:", search_account)
            for i in range(search_account):
                if paper_check(search_list, i, author_status, author_name, paper_name, mode=1):
                    p_c = search_list[i*7+1].get_attribute('innerText')
                    a_n = search_list[i*7+2].get_attribute('innerText')
                    p_t = search_list[i*7].get_attribute('innerText')
                    print(i, p_c, a_n, p_t)
                    access_index = i*7
                    break

            if access_index == -1:
                print("Can't find the paper by using paper name!")
                return []
            
            paper_href = paper_root_href + str((search_list[access_index].get_attribute('innerHTML').split("\""))[1])
            search_title = search_list[access_index].get_attribute('innerText')
            # Go to paper link
            paper_contents = get_paper_contents(chrome, paper_href)
            paper_contents[8][1] = search_title

            return paper_contents
    except:
        print("Can't find the paper by using paper name!")
        return []

################################################
# search_by_author_name
#   search the paper by author name, 
#   1. click the 論文名稱及研究生 circle button, and enter the author name.
#   2. Check the search result content including the paper name
#       if the count of result large than 100, use the school filter to find specific result.
#       else find and enter the paper content website.
################################################

def search_by_author_name(chrome, author_name, paper_name, author_status):
    print("Search the paper by author name....")
    chrome.get(    chrome.get(paper_root_href))
    try:
        chrome.find_element(By.ID, 'ti_論文名稱').click()
        chrome.find_element(By.ID, 'au_研究生').click()
        sleep(0.2)

        qs0 = chrome.find_element("id", "ysearchinput0")
        qs0.send_keys(author_name)
        qs0.submit()
        sleep(0.1)

        search = chrome.find_element(By.XPATH, '//td[@class="tdfmt1-content"]')
        sub_button = chrome.find_element(By.XPATH, '//div[@id="researchdivid"]/input')
        if search:
            search_list = chrome.find_elements(By.XPATH, '//td[@class="tdfmt1-content"]/div/div/table/tbody//td[@headers="simplefmt2td"]')
            search_account = int(len(search_list)/7)
            print(author_name, "查詢結果有:", search_account, "筆...")
            if search_account >= 100:
                print("透過學校名稱篩選...")
                research = chrome.find_element("id", "research")
                research.send_keys("國立臺灣科技大學")
                sub_button.click()
                sleep(0.2)
            
            search = chrome.find_element(By.XPATH, '//td[@class="tdfmt1-content"]')
            if search:            
                search_list = chrome.find_elements(By.XPATH, '//td[@class="tdfmt1-content"]/div/div/table/tbody//td[@headers="simplefmt2td"]')
                search_account = int(len(search_list)/7)
                access_index = -1
                print(author_name, "查詢結果有:", search_account, "筆...")
                for i in range(search_account):
                    if paper_check(search_list, i, author_status, author_name, paper_name, mode=1):
                        p_c = search_list[i*7+1].get_attribute('innerText')
                        a_n = search_list[i*7+2].get_attribute('innerText')
                        p_t = search_list[i*7].get_attribute('innerText')
                        print(i, p_c, a_n, p_t)
                        access_index = i*7
                        break
                
                if access_index < 0:
                    print("Can't find the paper by using paper author name!", access_index)
                    return []     

                print("存取第", int(access_index/7), "個")
                paper_href = paper_root_href + str((search_list[access_index].get_attribute('innerHTML').split("\""))[1])
                search_title = search_list[access_index].get_attribute('innerText')
                # Go to paper link
                paper_contents = get_paper_contents(chrome, paper_href)
                paper_contents[8][1] = search_title

                return paper_contents
    except:
        print("Can't find the paper by using paper author name!")
        return []

def do_paper_crawler(start, end, chrome):
    papers = read_excel(input_excel_name) # type(papers) = df Frame
    length = len(papers)
    print(length)
    papers['名稱可能有誤'] = ["" for i in range(length)]

    check_list = []
    
    for k in range(start, end+1):
        try:
            print("第 ", k, " 列讀取中...")
            i = k-2
            paper_name = papers['書名'][i]
            author_name = papers['作者'][i]
            author_status = papers['學位'][i]
            print(paper_name, author_name, author_status)

            # First search paper by author name, if not search then use paper name.
            result = []
            result = search_by_author_name(chrome, author_name, paper_name, author_status)
            if len(result) == 0:
                result= search_by_paper_name(chrome, author_name, paper_name, author_status)
            
            if len(result)!=0:
                print(result)
                for r in result:
                    papers[r[0]][i] = r[1]
                if papers["書名"][i] == papers["名稱可能有誤"][i]:
                    papers["名稱可能有誤"][i] = ""
            else:
                print("第", i+2, "列沒有找到相關論文")
        except:
            print("Error!")
    write_excel(papers, output_excel_name, sheet_name) 


if __name__ == "__main__":
    print("書目檔案:", input_excel_name)
    print("輸出檔案:", output_excel_name)
    start = 0

    try:
        while(start != -1):
            start = int(input("請輸入要查詢的起始列: "))
            end = int(input("請輸入要查詢的結束列: "))
            chrome = webdriver.Chrome(service=Service(ChromeDriverManager().install()), chrome_options=options)
            do_paper_crawler(start, end, chrome)
            print("第",start, "到", end, "列寫入完成囉~")
    except:
        print("程式終止")

