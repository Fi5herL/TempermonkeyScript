# TempermonkeyScript 🐵
TemperMonkey Scripts

- Active for [https://www.chlife-stat.org/](https://www.chlife-stat.org/) for MTG reports download.
   - 👤[ChurchLifeBasicInfoByNum.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeBasicInfoByNum.user.js) 匯出基本資料 (v4.3 - 配合VBA腳本版)
   - ⚙️[ChurchLifeYearByNumByMTG.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeYearByNumByMTG.user.js) 匯出各會所指定全年聚會資料
   - [ChurchLifeFullYearReportDownload.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeFullYearReportDownload.user.js) 匯出全年全部聚會資料
   - [ChurchLifeWeekReportCSVDownload.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeWeekReportCSVDownload.user.js) 下載CSV分頁點名檔(當周)
   - [ChurchLifeWeekReportCSVUploadRollCall.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeWeekReportCSVUploadRollCall.user.js) 上傳CSV自動點名(當周) ⚠️會修改到點名資訊
   - 取得 V1 點名系統 Cookie [ChurchLifeGetV1Cookie.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeGetV1Cookie.user.js)

- Active for Clock in/out system [Login-site](https://fa-eups-saasfaprod1.fa.ocs.oraclecloud.com/) for auto clock-in/out work days
   - [ClockInOutAutomation]()

- Active Project Price Summary [Flex](https://portal.ul.com/Dashboard)
   - [PjPriceSum.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/PjPriceSum.user.js)

## 報表資料彙整流程
   1. 依照Install Steps安裝油猴
   2. 安裝插件[ChurchLifeBasicInfoByNum.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeBasicInfoByNum.user.js) 與 [ChurchLifeYearByNumByMTG.user.js](https://github.com/Fi5herL/TempermonkeyScript/raw/refs/heads/main/ChurchLifeYearByNumByMTG.user.js)
   3. 登入[人數統計表網站](https://www.chlife-stat.org/)
   4. (聚會資料下載)進到 **"報表"** > **"聚會資料"** 頁面點選**⚙️**選擇要匯出的聚會選項
   5. 接著點選 **"匯出各會所項目(逐一)⚙️"**，等待檔案完全下載
   6. 下載完畢開一個新資料夾將所有聚會資料連同CSV目入移入
   7. (基本資料下載)接著進入 **"匯入/匯出"** 頁面點選 **開始自動匯出**，等待檔案完全下載
   8. 下載完畢開一個新資料夾將所有基本資料連同CSV目入移入
   9. (資料合併)請開啟[CHLifeMergeV2.2](https://github.com/Fi5herL/TempermonkeyScript/blob/main/CHLifeMergeV2.2(MTG%20BasicInfo%20and%20Verify%20Code%20Merge).xlsm)合併所有資料並儲存

## Install Steps 👟

---

1. Install browser extension [Tempermonkey](https://www.tampermonkey.net/)

   ![image](https://github.com/user-attachments/assets/1bcbc6f9-3ad2-463e-8bfb-8b14f3156bda)

---

2. Go to your browser **Manage Extension** setting and open **Developer mode**

   - Chrome: ``` chrome://extensions/ ```
   
   ![image](https://github.com/user-attachments/assets/4f470393-e217-436a-8b95-02cd18ba6f3c)

   - Edge: ``` edge://extensions/ ```

   ![image](https://github.com/user-attachments/assets/399ebbaf-b8eb-49c6-a976-68fae908caac)

---

3. Choose above script which you want, and click The **Raw** botton (Try [ChurchLifeFullYearReportDownload](https://github.com/Fi5herL/TempermonkeyScript/blob/main/ChurchLifeFullYearReportDownload.user.js) for example)

   ![image](https://github.com/user-attachments/assets/00098465-2c61-4a2d-b239-c1399334a873)

---

4. Wait a second and click **install** to apply the script

   ![image](https://github.com/user-attachments/assets/1c72f724-ca68-4f8b-a977-058c3c9adf14)

---

5. Go to your Target website and have fun

---

## Other Tools

- [操作紀錄器](https://greasyfork.org/zh-CN/scripts/461403-%E6%93%8D%E4%BD%9C%E8%AE%B0%E5%BD%95%E5%99%A8)
- [Chrome網頁分析助手](https://developer.chrome.com/docs/devtools/ai-assistance?hl=zh-tw)
