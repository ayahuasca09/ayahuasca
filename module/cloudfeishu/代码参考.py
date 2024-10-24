#!/usr/bin/env python
# coding=utf-8

import json
import requests
import os
import sys


class FeishuSheetHelper:
    def __init__(self):
        self.appID = "cli_a67dd70d527bd013"
        self.appSecret = "kw5O5VboETIgZERiNe39ebqL72kzbruu"
        self.SheetURL = ""
        self.tat = ""
        self.wikiSheetToken = ""

        self.SaveFileName = ""
        self.WriteRange = ""
        self.WriteDataList = ""
        self.SheetTitle = ""

    def RunFromCMD(self):
        mode = sys.argv[1]
        self.ParseArgs()
        self.InitTat()
        if mode == 'GetFirstSheetContent':
            self.GetFirstSheetContent(self.SheetURL, self.SaveFileName)
        if mode == 'GetSheetContent':
            self.GetSheetContent(self.SheetTitle, self.SheetURL, self.SaveFileName)
        elif mode == 'WriteToFirstSheetContent':
            self.WriteToFirstSheetContent(self.SheetURL, self.WriteRange, self.WriteDataList)

    def ParseArgs(self):
        for i in range(1, len(sys.argv)):
            if sys.argv[i] == "-SheetURL":
                self.SheetURL = sys.argv[i + 1]
            if sys.argv[i] == "-SaveFileName":
                self.SaveFileName = sys.argv[i + 1]
            if sys.argv[i] == "-WriteRange":
                self.WriteRange = sys.argv[i + 1]
            if sys.argv[i] == "-WriteDataList":
                self.WriteDataList = sys.argv[i + 1]
            if sys.argv[i] == "-SheetTitle":
                self.SheetTitle = sys.argv[i + 1]

    def GetFirstSheetContent(self, SheetURL="", SaveFileName=""):
        sheetToken = self.InitWikiSheetToken(SheetURL)
        sheets = self.GetWikiSheetInfos(sheetToken)  # 获取网页中所有表格页信息
        if len(sheets) > 0:
            sheetInfo = sheets[0]  # 获取第一个表格页信息
            jsonInfo = self.GetWikiSheetContentBySheetInfo(sheetInfo)
            self.SaveSheetContentToFile(jsonInfo, SaveFileName)
        else:
            print("Error: No sheet found in the wiki page")

    def GetSheetContent(self, SheetTitle, SheetURL="", SaveFileName=""):
        sheetToken = self.InitWikiSheetToken(SheetURL)
        sheets = self.GetWikiSheetInfos(sheetToken)  # 获取网页中所有表格页信息
        if len(sheets) > 0:
            sheetInfo = None
            for sheet in sheets:
                if sheet["title"] == SheetTitle:
                    sheetInfo = sheet

            if sheetInfo != None:
                jsonInfo = self.GetWikiSheetContentBySheetInfo(sheetInfo)
                self.SaveSheetContentToFile(jsonInfo, SaveFileName)
            else:
                print(f"Error: No sheet found in the wiki page with title {SheetTitle}")
        else:
            print("Error: No sheet found in the wiki page")

    def WriteToFirstSheetContent(self, SheetURL="", WriteRange="", WriteDataList=""):
        WriteRangeParts = WriteRange.split(",")
        if len(WriteRangeParts) != 4:
            print("Error: WriteRange is invalid")
            return False

        sheetToken = self.InitWikiSheetToken(SheetURL)
        sheets = self.GetWikiSheetInfos(sheetToken)
        if len(sheets) > 0:
            sheetInfo = sheets[0]
            sheetID = FeishuSheetHelper.GetSheetIDFromSheetInfo(sheetInfo)
            rangeStr = f"{sheetID}!{FeishuSheetHelper.GetCellStr(WriteRangeParts[0], WriteRangeParts[1])}:{FeishuSheetHelper.GetCellStr(WriteRangeParts[2], WriteRangeParts[3])}"
            data = FeishuSheetHelper.GetTwoDimensionalArrayFromString(WriteDataList)
            # print(rangeStr)
            # print(data)
            r = self.WriteWikiSheetContent(sheetID, rangeStr, data)
            print(r.json())
            if r.ok and r.json()["code"] == 0:
                print("Write data to sheet successfully")
            else:
                print(f"Error: Write data to sheet failed, {r.text}")
        else:
            print("Error: No sheet found in the wiki page")

    def GetAccessToken(self):
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
        post_data = {"app_id": self.appID, "app_secret": self.appSecret}
        r = requests.post(url, data=post_data)
        tat = r.json()["tenant_access_token"]
        # print("GetAccessToken: " + tat)
        return tat

    def InitTat(self):
        self.tat = self.GetAccessToken()

    def InitWikiSheetToken(self, WikiSheetURL):
        self.wikiSheetToken = self.GetWikiSheetToken(FeishuSheetHelper.GetWikiPageTokenFromURL(WikiSheetURL))
        return self.wikiSheetToken

    def GetWikiSheetToken(self, WikiPageToken, tat=""):
        if tat == "":
            tat = self.tat
        header = {"content-type": "application/json", "Authorization": "Bearer " + str(tat)}
        url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WikiPageToken}"
        r = requests.get(url, headers=header)
        # print(r.ok)
        if r.ok:
            # print(r.json())
            token = r.json()["data"]["node"]["obj_token"]
            return token
        else:
            return ""

    def GetWikiSheetInfos(self, WikiSheetToken="", tat=""):
        if tat == "":
            tat = self.tat
        if WikiSheetToken == "":
            WikiSheetToken = self.wikiSheetToken
        header = {"content-type": "application/json", "Authorization": "Bearer " + str(tat)}
        url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{WikiSheetToken}/sheets/query"
        r = requests.get(url, headers=header)
        print(r.json())
        if r.ok and r.json()["code"] == 0:
            # print(r.json())
            return r.json()["data"]["sheets"]
        else:
            return ""

    def GetWikiSheetContentBySheetInfo(self, sheetInfo, WikiSheetToken="", tat=""):
        sheetID = FeishuSheetHelper.GetSheetIDFromSheetInfo(sheetInfo)
        dataRange = FeishuSheetHelper.GetSheetRangeFromSheetInfo(sheetInfo)
        # print (sheetID, dataRange)
        return self.GetWikiSheetContent(sheetID, dataRange, WikiSheetToken, tat)

    def GetWikiSheetContent(self, sheetID, dataRange, WikiSheetToken="", tat=""):
        if tat == "":
            tat = self.tat
        if WikiSheetToken == "":
            WikiSheetToken = self.wikiSheetToken
        header = {"content-type": "application/json", "Authorization": "Bearer " + str(tat)}
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{WikiSheetToken}/values/{sheetID}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString"
        r = requests.get(url, headers=header)

        if r.ok:
            # print(r.json())
            return r.json()
        else:
            return ""

    def WriteWikiSheetContent(self, sheetID, dataRange, data, WikiSheetToken="", tat=""):
        if tat == "":
            tat = self.tat
        if WikiSheetToken == "":
            WikiSheetToken = self.wikiSheetToken
        header = {"content-type": "application/json", "Authorization": "Bearer " + str(tat)}
        payload = json.dumps({
            'valueRange': {
                'range': dataRange,
                'values': data
            }
        })
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{WikiSheetToken}/values"
        r = requests.put(url, headers=header, data=payload)
        return r

    def SaveSheetContentToFile(self, sheetContent, filePath):
        if 'data' in sheetContent:
            # Save two-dimensional array to a file, use "|" as separator
            with open(filePath, "w", encoding='utf-8') as f:
                values = sheetContent['data']['valueRange']['values']
                # print(values)
                for row in values:
                    strRow = [str(I) for I in row]
                    for i in range(len(strRow)):
                        strRow[i] = strRow[i].replace("\n", "")
                    f.write("|".join(strRow) + "\n")

            print(f"Sheet content has been saved to {filePath}")
        else:
            print("Error: Sheet content is invalid, cannot save to file")

    def GetWikiPageTokenFromURL(url):
        return url.split("/")[-1]

    def GetSheetIDFromSheetInfo(sheetInfo):
        return sheetInfo["sheet_id"]

    def GetSheetRangeFromSheetInfo(sheetInfo):
        print(sheetInfo)
        columnCount = sheetInfo["grid_properties"]["column_count"]
        rowCount = sheetInfo["grid_properties"]["row_count"]
        columnName = FeishuSheetHelper.ConvertIndexToExcelColumn(columnCount - 1)
        dataRange = f"A1:{columnName}{rowCount}"
        return dataRange

    def ConvertIndexToExcelColumn(index):
        if index < 26:
            return chr(65 + index)
        else:
            return chr(65 + index // 26 - 1) + chr(65 + index % 26)

    def GetCellStr(X, Y):
        X = int(X)
        Y = int(Y)
        return f"{FeishuSheetHelper.ConvertIndexToExcelColumn(X)}{Y}"

    # Parse two-dimensional array from string by '|' separator and ';' separator
    def GetTwoDimensionalArrayFromString(ArrayData):
        # print('ArrayData:', ArrayData)
        rows = ArrayData.split(";")
        result = []
        for row in rows:
            # print('row:', row)
            result.append(row.split("|"))
        return result


SheetHelper = FeishuSheetHelper()
SheetHelper.RunFromCMD()
