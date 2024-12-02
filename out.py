import base64
import hashlib
import math
import os
import time
from io import BytesIO
import ddddocr
from PIL import Image
from selenium import webdriver
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import yaml



pwd = os.getcwd()
configName = pwd + r"/config.yaml"

with open(configName, 'r', encoding='utf-8') as f:
    result = yaml.load(f.read(), Loader=yaml.FullLoader)

sign = result['sign']

# vars
chromeDriverPath = result["config"]['driver']
chromePath = result["config"]['browser']
baseurl = result["login"]["url"]
user = result["login"]["user"]
password = result["login"]["password"]
inFile = pwd + rf'/{result["config"]["in"]}'
outFile =pwd + rf'/{result["config"]["out"]}'

def signCheck(password, salt=None):
    salt = b"5qKF5bmy6I+c5bCP6YWl6aW8"
    password_with_salt = password.encode('utf-8') + salt
    sha512 = hashlib.sha512()
    sha512.update(password_with_salt)
    return sha512.hexdigest()

signCheckText = signCheck(str(user))
if signCheckText != sign:
    print("验签未通过，请重新添加验签值")
    time.sleep(10000000)
    exit()

df = pd.read_excel(outFile)
foodKeys = list(result["out"].keys())



# 判断
foodName = ""
for i in df.index:
    foodNameFromFile = df["名称"].loc[i]
    if "（" in foodNameFromFile:
        foodName = str(df["名称"].loc[i]).split("（")[0]
    elif " " in foodNameFromFile:
        foodName = str(df["名称"].loc[i]).split(' ')[0]
    else:
        foodName = str(foodNameFromFile)
    if foodName not in foodKeys:
        print(f'\n {foodNameFromFile} 不在配置文件中，请添加！, 本次将忽略此食品，请注意手动添加！')
        df = df[df["名称"] != foodNameFromFile]

df = df.reset_index(drop=True)
print()

# =================start
opt = Options()
opt.binary_locatio=fr"{chromePath}"
service = Service(r"{chromeDriverPath}")
driver = webdriver.Chrome()
driver.maximize_window()
# 打开指定网站
driver.get(baseurl)
time.sleep(3)
# 选择学校
driver.find_element(By.CSS_SELECTOR, '#tab-3').click()
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, 'input[placeholder="请输入用户名"]').send_keys(user)
driver.find_element(By.CSS_SELECTOR, 'input[placeholder="请输入密码"]').send_keys(password)
imgBase64 = driver.find_element(By.CSS_SELECTOR,'img[title="点击刷新"]').get_attribute("src").replace("data:image/gif;base64,", "")
veifyImg = BytesIO(base64.b64decode(imgBase64))
img = Image.open(veifyImg)
img.save("verify.png")
det = ddddocr.DdddOcr(show_ad=False)
verifypng = open(pwd+r"/verify.png", "rb").read()
res = det.classification(verifypng)
driver.find_element(By.CSS_SELECTOR, 'input[placeholder="请输入验证码"]').send_keys(res)
time.sleep(2)

# 登录
driver.find_element(By.CSS_SELECTOR, 'button span').click()
time.sleep(3)
# 进入采购管理
driver.find_element(By.XPATH, "//div[text()='库存管理']").click()
time.sleep(1)
# 出库
driver.find_element(By.CSS_SELECTOR, "#tab-export").click()
time.sleep(5)
driver.find_element(By.XPATH, "//span[text()='手动出库']").click()
time.sleep(2)


# 从excel表中获取数据
for j in df.index:
    if "（" in str(df["名称"].loc[j]):
        name = str(df["名称"].loc[j]).split("（")[0]
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').clear()
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').send_keys(name)
        driver.find_element(By.XPATH, '//form//form/div[2]/div/button/span').click()
    elif " " in str(df["名称"].loc[j]):
        name = str(df["名称"].loc[j]).split(' ')[0]
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').clear()
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').send_keys(name)
        driver.find_element(By.XPATH, '//form//form/div[2]/div/button/span').click()
    else:
        name = str(df["名称"].loc[j])
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').clear()
        driver.find_element(By.CSS_SELECTOR, 'input[placeholder="输入食材名称"]').send_keys(name)
        driver.find_element(By.XPATH, '//form//form/div[2]/div/button/span').click()
    # 总价
    if j > 0:
        p1 = f'//form//div/table/tbody/tr[{j+1}]/td[7]//div/div/input'
        p2 = f'//form//div/table/tbody/tr[{j+1}]/td[6]//div/div/input'
    else:
        p1 = '//form//div/table/tbody/tr/td[7]//div/div/input'
        p2 = '//form//div/table/tbody/tr/td[6]//div/div/input'

    time.sleep(2)
    pageNum = math.ceil(int(str(driver.find_element(By.CSS_SELECTOR, 'form span.el-pagination__total').text).replace("共", "").replace("条", "")) / 15)
    for sec in range(pageNum):
        flag = False
        for i in range(1, 16):
            # 分类
            categorySelector = f"form div.el-table__body-wrapper  tbody tr.el-table__row:nth-child({i}) td:nth-child(1)"
            # 食材名称
            foodNameSelector = f"form div.el-table__body-wrapper  tbody tr.el-table__row:nth-child({i}) td:nth-child(2)"
            # 规格
            specSelector = f"form div.el-table__body-wrapper  tbody tr.el-table__row:nth-child({i}) td:nth-child(4)"
            # 单位
            unitSelector = f"form div.el-table__body-wrapper  tbody tr.el-table__row:nth-child({i}) td:nth-child(5)"
            # 添加
            addSelector = f"form div.el-table__body-wrapper  tbody tr.el-table__row:nth-child({i}) td:nth-child(6)"
            check1Text = driver.find_element(By.CSS_SELECTOR, categorySelector).text
            check2Text = driver.find_element(By.CSS_SELECTOR, foodNameSelector).text
            check3Text = driver.find_element(By.CSS_SELECTOR, specSelector).text
            check4Text = driver.find_element(By.CSS_SELECTOR, unitSelector).text
            isClick = False
            # 是否需要特殊处理
            checkNum = len(result["out"][name].keys())
            for i in range(checkNum):
                temp = list(result["out"][name].keys())[i]
                if temp == "分类":
                    check1 = result["out"][name][temp]
                    if check1 == check1Text:
                        isClick = True
                    else:
                        isClick = False
                if temp == "名称":
                    check2 = result["out"][name][temp]
                    if check2 == check2Text:
                        isClick = True
                    else:
                        isClick = False
                if temp == "规格":
                    check3 = result["out"][name][temp]
                    if check3 == check3Text:
                        isClick = True
                    else:
                        isClick = False
                if temp == "单位":
                    check4 = result["out"][name][temp]
                    if check4 == check4Text:
                        isClick = True
                    else:
                        isClick = False

            if isClick == True:
                # 判断正确
                driver.find_element(By.CSS_SELECTOR, addSelector).click()
                driver.find_element(By.XPATH, p1).clear()
                driver.find_element(By.XPATH, p1).send_keys(str(df["出库数量"].loc[j]))
                driver.find_element(By.XPATH, p1).send_keys(Keys.TAB)
                time.sleep(2)
                flag = True
                print("已经成功添加出库信息:         ", name)
                break
        # 下一页
        if flag == True:
            break
        nextSelector = "form button.btn-next"
        driver.find_element(By.CSS_SELECTOR, nextSelector).click()
        time.sleep(1)



time.sleep(90000)
