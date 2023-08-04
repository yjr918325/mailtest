from xlrd import open_workbook
import re
from xlutils.copy import copy
import datetime
import xlwt


def readExcel(filename, n):
    excel = open_workbook(filename)
    sheet = excel.sheets()[0]
    col = sheet.col_values(n)  # n=12:waring content ; n=44:state of kaidan
    return col


# excel未处理不需要的告警！
# 归纳合并所需的告警类信息
def extractData(col):
    bracket = re.compile('(\()')
    idPattern = re.compile(r'(\()(.*?)(\))')
    timePattern = re.compile('.* ([0-9][0-9]):([0-9][0-9]).*')
    MON_OBJ_ID = []
    timeHour = []
    p1Dict = {}
    # p2Dict = {}
    count = 1
    for p in col[1:]:
        errorList = []
        if "秒内成功率" in p:
            errorList.append("succRate")
        if "秒无交易上送" in p:
            errorList.append("noTransTm")
        if "交易量下降" in p or "交易下降了" in p:
            errorList.append("fluctValue")
        if "秒内失败笔数" in p:
            errorList.append("transFailNum")
        if "连续失败笔数" in p:
            errorList.append("seqFailNum")
        if "ERROR_CD" in p:
            errorList.append("httpCnt")
        try:
            objId = idPattern.search(p).group(2)
            while len(objId) < 5:
                ret = bracket.sub('', p, 1)
                objId = idPattern.search(ret).group(2)
                p = ret
            objTime = str(timePattern.search(p).group(1))
            # print(objTime)
            MON_OBJ_ID.append(objId)
            timeHour.append(objTime)
            # print(objTime,errorList,count)
        except:
            print("illegal data format.")
        if objId not in p1Dict:
            p1Dict[objId] = {}
        # print(errorList,objId)
        for ruleType in errorList:
            if ruleType not in p1Dict[objId]:
                p1Dict[objId][ruleType] = []
            if len(objTime) == 2 and objTime[0] == '0':
                objTime = objTime[1]
            p1Dict[objId][ruleType].append(objTime)
            # print(p2Dict,count,objId)
        # p1Dict[objId]=p2Dict
    print(p1Dict)
    return p1Dict


# 写入auto_rule文件
def writeExcelRule(p1Dict):
    nowTime = datetime.datetime.now().strftime("%m%d")
    fileName = "auto_rule_conf_" + nowTime + '.xls'
    workBook = xlwt.Workbook(encoding='utf-8')
    sheet = workBook.add_sheet('规则调优')
    sheet.write(0, 0, 'MON_OBJ_ID')
    sheet.write(0, 1, 'DATA_SOURCE_ID')
    sheet.write(0, 2, 'RULE_TYPE')
    sheet.write(0, 3, 'METHOD')
    sheet.write(0, 4, 'TIME')
    sheet.write(0, 5, 'MULTI_TRIG')
    sheet.write(0, 6, 'Frequency')
    workBook.save(fileName)
    autoRuleNum = 1
    for id in p1Dict:
        pattern = re.compile('(.*?)_')
        try:
            sourceId = pattern.match(id).group(1)
        except:
            print('Data fail')
            continue
        for errorType in p1Dict[id]:
            oldbook = open_workbook(fileName)
            wb = copy(oldbook)
            ws = wb.get_sheet(0)
            ws.write(autoRuleNum, 0, id)
            ws.write(autoRuleNum, 1, sourceId)
            ws.write(autoRuleNum, 2, errorType)
            ws.write(autoRuleNum, 3, 'auto')
            # 写入符合规格的时间信息
            seq = len(p1Dict[id][errorType])
            set1 = set(p1Dict[id][errorType])
            setlist = list(set1)
            setlist.sort()
            timeLog = ",".join(setlist)
            ws.write(autoRuleNum, 4, timeLog)
            ws.write(autoRuleNum, 5, '1')
            ws.write(autoRuleNum, 6, str(seq))
            wb.save(fileName)
            autoRuleNum += 1

    #         autoRuleNum+=1
    # for row in range(1,autoRuleNum+1):
    #     oldbook = open_workbook(fileName)
    #     wb = copy(oldbook)
    #     ws = wb.get_sheet(0)


# 写入告警分析模板
def writeExcelAnalyse(col_content, col_event):
    UPACP_All_Warning = UPACP_All_Event = UPACP_All_object = 0
    UPACP_All_object_list = []
    UPACP_Type_Warning = UPACP_Type_Event = UPACP_Type_object = 0
    UPACP_Type_object_list = []
    UPACP_procsys_Warning = UPACP_procsys_Event = UPACP_procsys_object = 0
    UPACP_procsys_object_list = []
    UPACP_client_id_Warning = UPACP_client_id_Event = UPACP_client_id_object = 0
    UPACP_client_id_object_list = []
    UPACP_Merchant_Warning = UPACP_Merchant_Event = UPACP_Merchant_object = 0
    UPACP_Merchant_object_list = []
    HTTP_UPACP_Warning = HTTP_UPACP_Event = HTTP_UPACP_object = 0
    HTTP_UPACP_object_list = []
    HTTP_UPACP_qzhlw_Warning = HTTP_UPACP_qzhlw_Event = HTTP_UPACP_qzhlw_object = 0
    HTTP_UPACP_qzhlw_object_list = []

    QRC_QRC_Warning = QRC_QRC_Event = QRC_QRC_object = 0
    QRC_QRC_object_list = []
    QRC_Center_Warning = QRC_Center_Event = QRC_Center_object = 0
    QRC_Center_object_list = []
    QRC_Type_Warning = QRC_Type_Event = QRC_Type_object = 0
    QRC_Type_object_list = []
    QRC_Order_Warning = QRC_Order_Event = QRC_Order_object = 0
    QRC_Order_object_list = []
    HTTP_QRC_All_Warning = HTTP_QRC_All_Event = HTTP_QRC_All_object = 0
    HTTP_QRC_All_object_list = []
    HTTP_QRC_zxhlw_Warning = HTTP_QRC_zxhlw_Event = HTTP_QRC_zxhlw_object = 0
    HTTP_QRC_zxhlw_object_list = []
    HTTP_QRC_error_Warning = HTTP_QRC_error_Event = HTTP_QRC_error_object = 0
    HTTP_QRC_error_object_list = []

    FS_FS_waring = FS_FS_Event = FS_FS_object = 0
    FS_FS_object_list = []
    FS_Center_waring = FS_Center_Event = FS_Center_object = 0
    FS_Center_object_list = []
    FS_Type_waring = FS_Type_Event = FS_Type_object = 0
    FS_Type_object_list = []
    FS_procSys_waring = FS_procSys_Event = FS_procSys_object = 0
    FS_procSys_object_list = []
    HTTP_fsasAll_waring = HTTP_fsasAll_Event = HTTP_fsasAll_object = 0
    HTTP_fsasAll_object_list = []
    HTTP_FS_qzhlw_waring = HTTP_FS_qzhlw_Event = HTTP_FS_qzhlw_object = 0
    HTTP_FS_qzhlw_object_list = []
    HTTP_FS_error_waring = HTTP_FS_error_Event = HTTP_FS_error_object = 0
    HTTP_FS_error_object_list = []

    QTT_QTTB_waring = QTT_QTTB_Event = QTT_QTTB_object = 0
    QTT_QTTB_object_list = []
    QTT_Center_waring = QTT_Center_Event = QTT_Center_object = 0
    QTT_Center_object_list = []
    QTT_Type_waring = QTT_Type_Event = QTT_Type_object = 0
    QTT_Type_object_list = []
    HTTP_QTTall_waring = HTTP_QTTall_Event = HTTP_QTTall_object = 0
    HTTP_QTTall_object_list = []
    HTTP_QTTdomain_waring = HTTP_QTTdomain_Event = HTTP_QTTdomain_object = 0
    HTTP_QTTdomain_object_list = []
    HTTP_QTThlwqz_waring = HTTP_QTThlwqz_Event = HTTP_QTThlwqz_object = 0
    HTTP_QTThlwqz_object_list = []
    HTTP_QTThlw_qz_waring = HTTP_QTThlw_qz_Event = HTTP_QTThlw_qz_object = 0
    HTTP_QTThlw_qz_object_list = []
    HTTP_QTT_error_waring = HTTP_QTT_error_Event = HTTP_QTT_error_object = 0
    HTTP_QTT_error_object_list = []

    QAT_QAT_waring = QAT_QAT_Event = QAT_QAT_object = 0
    QAT_QAT_object_list = []
    QAT_Center_waring = QAT_Center_Event = QAT_Center_object = 0
    QAT_Center_object_list = []
    QAT_Type_waring = QAT_Type_Event = QAT_Type_object = 0
    QAT_Type_object_list = []
    HTTP_QATall_waring = HTTP_QATall_Event = HTTP_QATall_object = 0
    HTTP_QATall_object_list = []
    HTTP_QATdomain_waring = HTTP_QATdomain_Event = HTTP_QATdomain_object = 0
    HTTP_QATdomain_object_list = []
    HTTP_QATAPI_waring = HTTP_QATAPI_Event = HTTP_QATAPI_object = 0
    HTTP_QATAPI_object_list = []
    HTTP_QAT_error_waring = HTTP_QAT_error_Event = HTTP_QAT_error_object = 0
    HTTP_QAT_error_object_list = []

    idPattern = re.compile(r'(\()(.*?)(\))')
    bracket = re.compile('(\()')
    index = 1

    # 按id归类，待补充
    for content in col_content[1:]:
        try:
            objId = idPattern.search(content).group(2)
            while len(objId) < 5:
                ret = bracket.sub('', content, 1)
                objId = idPattern.search(ret).group(2)
                content = ret
        except:
            print("illegal data format.")
            continue
        if objId in ['UPACPNOTIFY_UPACPNOTIFY_NA', 'UPACP_UPACP_NA', 'UPACPTRANS_UPACPTRANS_NA']:
            UPACP_All_Warning += 1
            if col_event[index] == '已开单':
                UPACP_All_Event += 1
            UPACP_All_object_list.append(objId)
        elif 'UPACP' in objId and 'Trans' in objId:
            UPACP_Type_Warning += 1
            if col_event[index] == '已开单':
                UPACP_Type_Event += 1
            UPACP_Type_object_list.append(objId)
        elif 'UPACP' in objId and 'PS' in objId:
            UPACP_procsys_Warning += 1
            if col_event[index] == '已开单':
                UPACP_procsys_Event += 1
            UPACP_procsys_object_list.append(objId)
        elif 'UPACP_ClientID' in objId:
            UPACP_client_id_Warning += 1
            if col_event[index] == '已开单':
                UPACP_client_id_Event += 1
            UPACP_client_id_object_list.append(objId)
        elif 'UPACP_' in objId and '_M' in objId:
            UPACP_Merchant_Warning += 1
            if col_event[index] == '已开单':
                UPACP_Merchant_Event += 1
            UPACP_Merchant_object_list.append(objId)
        elif objId == 'HTTP_UPACP_SYSTEM':
            HTTP_UPACP_Warning += 1
            if col_event[index] == '已开单':
                HTTP_UPACP_Event += 1
            HTTP_UPACP_object_list.append(objId)
        elif objId in ['HTTP_upacqz_hlw', 'HTTP_upacqz_zx']:
            HTTP_UPACP_qzhlw_Warning += 1
            if col_event[index] == '已开单':
                HTTP_UPACP_qzhlw_Event += 1
            HTTP_UPACP_qzhlw_object_list.append(objId)
        elif objId == "QRC_QRC_NA":
            QRC_QRC_Warning += 1
            if col_event[index] == '已开单':
                QRC_QRC_Event += 1
            QRC_QRC_object_list.append(objId)
        elif objId in ['QRC_U_NA', 'QRC_H_NA']:
            QRC_Center_Warning += 1
            if col_event[index] == '已开单':
                QRC_Center_Event += 1
            QRC_Center_object_list.append(objId)
        elif 'QRC_QRCUP' in objId or objId == 'QRC_01_NA':
            QRC_Type_Warning += 1
            if col_event[index] == '已开单':
                QRC_Type_Event += 1
            QRC_Type_object_list.append(objId)
        elif 'QRC_' in objId and '_U_NA' in objId:
            QRC_Order_Warning += 1
            if col_event[index] == '已开单':
                QRC_Order_Event += 1
            QRC_Order_object_list.append(objId)
        elif 'QRC_' in objId and '_H_NA' in objId:
            QRC_Order_Warning += 1
            if col_event[index] == '已开单':
                QRC_Order_Event += 1
            QRC_Order_object_list.append(objId)
        elif objId == 'HTTP_qrc_system':
            HTTP_QRC_All_Warning += 1
            if col_event[index] == '已开单':
                HTTP_QRC_All_Event += 1
            HTTP_QRC_All_object_list.append(objId)
        elif objId in ['HTTP_qrc_zhuanxian_000001', 'HTTP_qrc_hlw_000001']:
            HTTP_QRC_zxhlw_Warning += 1
            if col_event[index] == '已开单':
                HTTP_QRC_zxhlw_Event += 1
            HTTP_QRC_zxhlw_object_list.append(objId)
        elif 'HTTP_qrc_system_ERROR_CD' in objId:
            HTTP_QRC_error_Warning += 1
            if col_event[index] == '已开单':
                HTTP_QRC_error_Event += 1
            HTTP_QRC_error_object_list.append(objId)
        elif objId == 'FS_FS_NA':
            FS_FS_waring += 1
            if col_event[index] == '已开单':
                FS_FS_Event += 1
            FS_FS_object_list.append(objId)
        elif 'FS_CENTER_Center' in objId:
            FS_Center_waring += 1
            if col_event[index] == '已开单':
                FS_Center_Event += 1
            FS_Center_object_list.append(objId)
        elif 'FS_TYPE_' in objId:
            FS_Type_waring += 1
            if col_event[index] == '已开单':
                FS_Type_Event += 1
            FS_Type_object_list.append(objId)
        elif 'FS_PROC_SYS_' in objId:
            FS_procSys_waring += 1
            if col_event[index] == '已开单':
                FS_procSys_Event += 1
            FS_procSys_object_list.append(objId)
        elif objId == 'HTTP_fsas_system':
            HTTP_fsasAll_waring += 1
            if col_event[index] == '已开单':
                HTTP_fsasAll_Event += 1
            HTTP_fsasAll_object_list.append(objId)
        elif objId in ['HTTP_fsas_zx', 'HTTP_fsas_hlw']:
            HTTP_FS_qzhlw_waring += 1
            if col_event[index] == '已开单':
                HTTP_FS_qzhlw_Event += 1
            HTTP_FS_qzhlw_object_list.append(objId)
        elif 'HTTP_fsas_system_ERROR_CD' in objId:
            HTTP_FS_error_waring += 1
            if col_event[index] == '已开单':
                HTTP_FS_error_Event += 1
            HTTP_FS_error_object_list.append(objId)
        elif objId == 'QTT_QTTB_B':
            QTT_QTTB_waring += 1
            if col_event[index] == '已开单':
                QTT_QTTB_Event += 1
            QTT_QTTB_object_list.append(objId)
        elif 'QTT_CENTER_S' in objId:
            QTT_Center_waring += 1
            if col_event[index] == '已开单':
                QTT_Center_Event += 1
            QTT_Center_object_list.append(objId)
        elif objId in ['QTT_JSAPI_B', 'QTT_MICROPAY_B', 'QTT_NATIVE_B', 'QTT_APP_B']:
            QTT_Type_waring += 1
            if col_event[index] == '已开单':
                QTT_Type_Event += 1
            QTT_Type_object_list.append(objId)
        elif objId == 'HTTP_T_System':
            HTTP_QTTall_waring += 1
            if col_event[index] == '已开单':
                HTTP_QTTall_Event += 1
            HTTP_QTTall_object_list.append(objId)
        elif objId == 'HTTP_QTT_DOMAINM_':
            HTTP_QTTdomain_waring += 1
            if col_event[index] == '已开单':
                HTTP_QTTdomain_Event += 1
            HTTP_QTTdomain_object_list.append(objId)
        elif objId in ['HTTP_T_QZ_ZX', 'HTTP_T_QZ_HLW']:
            HTTP_QTThlwqz_waring += 1
            if col_event[index] == '已开单':
                HTTP_QTThlwqz_Event += 1
            HTTP_QTThlwqz_object_list.append(objId)
        elif 'HTTP_T_QZ_ZX_' in objId or 'HTTP_T_QZ_HLW_' in objId:
            HTTP_QTThlw_qz_waring += 1
            if col_event[index] == '已开单':
                HTTP_QTThlw_qz_Event += 1
            HTTP_QTThlw_qz_object_list.append(objId)
        elif 'HTTP_TSystem_ERROR_CD_' in objId:
            HTTP_QTT_error_waring += 1
            if col_event[index] == '已开单':
                HTTP_QTT_error_Event += 1
            HTTP_QTT_error_object_list.append(objId)
        elif objId == 'QAT_QAT_NA':
            QAT_QAT_waring += 1
            if col_event[index] == '已开单':
                QAT_QAT_Event += 1
            QAT_QAT_object_list.append(objId)
        elif objId == 'QAT_CENTER_SH':
            QAT_Center_waring += 1
            if col_event[index] == '已开单':
                QAT_Center_Event += 1
            QAT_Center_object_list.append(objId)
        elif objId in ['QAT_pay_B', 'QAT_create_B', 'QAT_precreate_B']:
            QAT_Type_waring += 1
            if col_event[index] == '已开单':
                QAT_Type_Event += 1
            QAT_Type_object_list.append(objId)
        elif objId == 'HTTP_A_System':
            HTTP_QATall_waring += 1
            if col_event[index] == '已开单':
                HTTP_QATall_Event += 1
            HTTP_QATall_object_list.append(objId)
        elif objId in ['HTTP_A_QZ_ZX', 'HTTP_A_QZ_HLW']:
            HTTP_QATdomain_waring += 1
            if col_event[index] == '已开单':
                HTTP_QATdomain_Event += 1
            HTTP_QATdomain_object_list.append(objId)
        elif objId not in ['HTTP_A_QZ_ZX', 'HTTP_A_QZ_HLW'] and 'HTTP_A_QZ_' in objId:
            HTTP_QATAPI_waring += 1
            if col_event[index] == '已开单':
                HTTP_QATAPI_Event += 1
            HTTP_QATAPI_object_list.append(objId)
        elif 'HTTP_ASystem_ERROR_CD_' in objId:
            HTTP_QAT_error_waring += 1
            if col_event[index] == '已开单':
                HTTP_QAT_error_Event += 1
            HTTP_QAT_error_object_list.append(objId)
        index += 1

    UPACP_All_object = len(set(UPACP_All_object_list))
    UPACP_All_object_str = ",".join(set(UPACP_All_object_list))
    UPACP_Type_object = len(set(UPACP_Type_object_list))
    UPACP_Type_object_str = ",".join(set(UPACP_All_object_list))
    UPACP_procsys_object = len(set(UPACP_procsys_object_list))
    UPACP_procsys_object_str = ",".join(set(UPACP_procsys_object_list))
    UPACP_client_id_object = len(set(UPACP_client_id_object_list))
    UPACP_client_id_object_str = ",".join(set(UPACP_client_id_object_list))
    UPACP_Merchant_object = len(set(UPACP_Merchant_object_list))
    UPACP_Merchant_object_str = ",".join(set(UPACP_Merchant_object_list))
    HTTP_UPACP_object = len(set(HTTP_UPACP_object_list))
    HTTP_UPACP_object_str = ",".join(set(HTTP_UPACP_object_list))
    HTTP_UPACP_qzhlw_object = len(set(HTTP_UPACP_qzhlw_object_list))
    HTTP_UPACP_qzhlw_object_str = ",".join(set(HTTP_UPACP_qzhlw_object_list))

    QRC_QRC_object = len(set(QRC_QRC_object_list))
    QRC_QRC_object_str = ",".join(set(QRC_QRC_object_list))
    QRC_Center_object = len(set(QRC_Center_object_list))
    QRC_Center_object_str = ",".join(set(QRC_Center_object_list))
    QRC_Type_object = len(set(QRC_Type_object_list))
    QRC_Type_object_str = ",".join(set(QRC_Type_object_list))
    QRC_Order_object = len(set(QRC_Order_object_list))
    QRC_Order_object_str = ",".join(set(QRC_Order_object_list))
    HTTP_QRC_All_object = len(set(HTTP_QRC_All_object_list))
    HTTP_QRC_All_object_str = ",".join(set(HTTP_QRC_All_object_list))
    HTTP_QRC_zxhlw_object = len(set(HTTP_QRC_zxhlw_object_list))
    HTTP_QRC_zxhlw_object_str = ",".join(set(HTTP_QRC_zxhlw_object_list))
    HTTP_QRC_error_object = len(set(HTTP_QRC_error_object_list))
    HTTP_QRC_error_object_str = ",".join(set(HTTP_QRC_error_object_list))

    FS_FS_object = len(set(FS_FS_object_list))
    FS_FS_object_str = ",".join(set(FS_FS_object_list))
    FS_Center_object = len(set(FS_Center_object_list))
    FS_Center_object_str = ",".join(set(FS_Center_object_list))
    FS_Type_object = len(set(FS_Type_object_list))
    FS_Type_object_str = ",".join(set(FS_Type_object_list))
    FS_procSys_object = len(set(FS_procSys_object_list))
    FS_procSys_object_str = ",".join(set(FS_procSys_object_list))
    HTTP_fsasAll_object = len(set(HTTP_fsasAll_object_list))
    HTTP_fsasAll_object_str = ",".join(set(HTTP_fsasAll_object_list))
    HTTP_FS_qzhlw_object = len(set(HTTP_FS_qzhlw_object_list))
    HTTP_FS_qzhlw_object_str = ",".join(set(FS_FS_object_list))
    HTTP_FS_error_object = len(set(HTTP_FS_error_object_list))
    HTTP_FS_error_object_str = ",".join(set(HTTP_FS_error_object_list))

    QTT_QTTB_object = len(set(QTT_QTTB_object_list))
    QTT_QTTB_object_str = ",".join(set(QTT_QTTB_object_list))
    QTT_Center_object = len(set(QTT_Center_object_list))
    QTT_Center_object_str = ",".join(set(QTT_Center_object_list))
    QTT_Type_object = len(set(QTT_Type_object_list))
    QTT_Type_object_str = ",".join(set(QTT_Type_object_list))
    HTTP_QTTall_object = len(set(HTTP_QTTall_object_list))
    HTTP_QTTall_object_str = ",".join(set(HTTP_QTTall_object_list))
    HTTP_QTTdomain_object = len(set(HTTP_QTTdomain_object_list))
    HTTP_QTTdomain_object_str = ",".join(set(HTTP_QTTdomain_object_list))
    HTTP_QTThlwqz_object = len(set(HTTP_QTThlwqz_object_list))
    HTTP_QTThlwqz_object_str = ",".join(set(HTTP_QTThlwqz_object_list))
    HTTP_QTThlw_qz_object = len(set(HTTP_QTThlw_qz_object_list))
    HTTP_QTThlw_qz_object_str = ",".join(set(HTTP_QTThlw_qz_object_list))
    HTTP_QTT_error_object = len(set(HTTP_QTT_error_object_list))
    HTTP_QTT_error_object_str = ",".join(set(HTTP_QTT_error_object_list))

    QAT_QAT_object = len(set(QAT_QAT_object_list))
    QAT_QAT_object_str = ",".join(set(QAT_QAT_object_list))
    QAT_Center_object = len(set(QAT_Center_object_list))
    QAT_Center_object_str = ",".join(set(QAT_Center_object_list))
    QAT_Type_object = len(set(QAT_Type_object_list))
    QAT_Type_object_str = ",".join(set(QAT_Type_object_list))
    HTTP_QATall_object = len(set(HTTP_QATall_object_list))
    HTTP_QATall_object_str = ",".join(set(HTTP_QATall_object_list))
    HTTP_QATdomain_object = len(set(HTTP_QATdomain_object_list))
    HTTP_QATdomain_object_str = ",".join(set(HTTP_QATdomain_object_list))
    HTTP_QATAPI_object = len(set(HTTP_QATAPI_object_list))
    HTTP_QATAPI_object_str = ",".join(set(HTTP_QATAPI_object_list))
    HTTP_QAT_error_object = len(set(HTTP_QAT_error_object_list))
    HTTP_QAT_error_object_str = ",".join(set(HTTP_QAT_error_object_list))

    # 写入excel特定位置
    fileName = "告警分析1.xls"
    oldbook = open_workbook(fileName)
    wb = copy(oldbook)
    ws = wb.get_sheet(0)
    ws.write(1, 6, UPACP_All_Warning)
    ws.write(1, 8, UPACP_All_Event)
    ws.write(1, 9, UPACP_All_object)

    ws.write(3, 6, UPACP_Type_Warning)
    ws.write(3, 8, UPACP_Type_Event)
    ws.write(3, 9, UPACP_Type_object)

    ws.write(4, 6, UPACP_procsys_Warning)
    ws.write(4, 8, UPACP_procsys_Event)
    ws.write(4, 9, UPACP_procsys_object)

    ws.write(5, 6, UPACP_client_id_Warning)
    ws.write(5, 8, UPACP_client_id_Event)
    ws.write(5, 9, UPACP_client_id_object)

    ws.write(6, 6, UPACP_Merchant_Warning)
    ws.write(6, 8, UPACP_Merchant_Event)
    ws.write(6, 9, UPACP_Merchant_object)

    ws.write(8, 6, HTTP_UPACP_Warning)
    ws.write(8, 8, HTTP_UPACP_Event)
    ws.write(8, 9, HTTP_UPACP_object)

    ws.write(10, 6, HTTP_UPACP_qzhlw_Warning)
    ws.write(10, 8, HTTP_UPACP_qzhlw_Event)
    ws.write(10, 9, HTTP_UPACP_qzhlw_object)

    ws.write(13, 6, QRC_QRC_Warning)
    ws.write(13, 8, QRC_QRC_Event)
    ws.write(13, 9, QRC_QRC_object)

    ws.write(14, 6, QRC_Center_Warning)
    ws.write(14, 8, QRC_Center_Event)
    ws.write(14, 9, QRC_Center_object)

    ws.write(15, 6, QRC_Type_Warning)
    ws.write(15, 8, QRC_Type_Event)
    ws.write(15, 9, QRC_Type_object)

    ws.write(16, 6, QRC_Order_Warning)
    ws.write(16, 8, QRC_Order_Event)
    ws.write(16, 9, QRC_Order_object)

    ws.write(20, 6, HTTP_QRC_All_Warning)
    ws.write(20, 8, HTTP_QRC_All_Event)
    ws.write(20, 9, HTTP_QRC_All_object)

    ws.write(22, 6, HTTP_QRC_zxhlw_Warning)
    ws.write(22, 8, HTTP_QRC_zxhlw_Event)
    ws.write(22, 9, HTTP_QRC_zxhlw_object)

    ws.write(24, 6, HTTP_QRC_error_Warning)
    ws.write(24, 8, HTTP_QRC_error_Event)
    ws.write(24, 9, HTTP_QRC_error_object)

    ws.write(25, 6, FS_FS_waring)
    ws.write(25, 8, FS_FS_Event)
    ws.write(25, 9, FS_FS_object)

    ws.write(26, 6, FS_Center_waring)
    ws.write(26, 8, FS_Center_Event)
    ws.write(26, 9, FS_Center_object)

    ws.write(27, 6, FS_Type_waring)
    ws.write(27, 8, FS_Type_Event)
    ws.write(27, 9, FS_Type_object)

    ws.write(28, 6, FS_procSys_waring)
    ws.write(28, 8, FS_procSys_Event)
    ws.write(28, 9, FS_procSys_object)

    ws.write(30, 6, HTTP_fsasAll_waring)
    ws.write(30, 8, HTTP_fsasAll_Event)
    ws.write(30, 9, HTTP_fsasAll_object)

    ws.write(32, 6, HTTP_FS_qzhlw_waring)
    ws.write(32, 8, HTTP_FS_qzhlw_Event)
    ws.write(32, 9, HTTP_FS_qzhlw_object)

    ws.write(34, 6, HTTP_FS_error_waring)
    ws.write(34, 8, HTTP_FS_error_Event)
    ws.write(34, 9, HTTP_FS_error_object)

    ws.write(35, 6, QTT_QTTB_waring)
    ws.write(35, 8, QTT_QTTB_Event)
    ws.write(35, 9, QTT_QTTB_object)

    ws.write(36, 6, QTT_Center_waring)
    ws.write(36, 8, QTT_Center_Event)
    ws.write(36, 9, QTT_Center_object)

    ws.write(37, 6, QTT_Type_waring)
    ws.write(37, 8, QTT_Type_Event)
    ws.write(37, 9, QTT_Type_object)

    ws.write(39, 6, HTTP_QTTall_waring)
    ws.write(39, 8, HTTP_QTTall_Event)
    ws.write(39, 9, HTTP_QTTall_object)

    ws.write(40, 6, HTTP_QTTdomain_waring)
    ws.write(40, 8, HTTP_QTTdomain_Event)
    ws.write(40, 9, HTTP_QTTdomain_object)

    ws.write(41, 6, HTTP_QTThlwqz_waring)
    ws.write(41, 8, HTTP_QTThlwqz_Event)
    ws.write(41, 9, HTTP_QTThlwqz_object)

    ws.write(42, 6, HTTP_QTThlw_qz_waring)
    ws.write(42, 8, HTTP_QTThlw_qz_Event)
    ws.write(42, 9, HTTP_QTThlw_qz_object)

    ws.write(43, 6, HTTP_QTT_error_waring)
    ws.write(43, 8, HTTP_QTT_error_Event)
    ws.write(43, 9, HTTP_QTT_error_object)

    ws.write(44, 6, QAT_QAT_waring)
    ws.write(44, 8, QAT_QAT_Event)
    ws.write(44, 9, QAT_QAT_object)

    ws.write(45, 6, QAT_Center_waring)
    ws.write(45, 8, QAT_Center_Event)
    ws.write(45, 9, QAT_Center_object)

    ws.write(46, 6, QAT_Type_waring)
    ws.write(46, 8, QAT_Type_Event)
    ws.write(46, 9, QAT_Type_object)

    ws.write(48, 6, HTTP_QATall_waring)
    ws.write(48, 8, HTTP_QATall_Event)
    ws.write(48, 9, HTTP_QATall_object)

    ws.write(50, 6, HTTP_QATdomain_waring)
    ws.write(50, 8, HTTP_QATdomain_Event)
    ws.write(50, 9, HTTP_QATdomain_object)

    ws.write(51, 6, HTTP_QATAPI_waring)
    ws.write(51, 8, HTTP_QATAPI_Event)
    ws.write(51, 9, HTTP_QATAPI_object)

    ws.write(52, 6, HTTP_QAT_error_waring)
    ws.write(52, 8, HTTP_QAT_error_Event)
    ws.write(52, 9, HTTP_QAT_error_object)
#-----------------------------------写入sheet2----------------------------------------
    ws1 = wb.get_sheet(1)
    row = 1
    for i in set(UPACP_All_object_list):
        ws1.write(row,0,'全渠道')
        ws1.write(row,1,'UPACP')
        ws1.write(row,2,'整体对象监控')
        ws1.write(row,3,i)
        ws1.write(row,4,str(UPACP_All_object_list.count(i)))
        row += 1
    for i in set(UPACP_Type_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '交易类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(UPACP_Type_object_list.count(i)))
        row += 1
    for i in set(UPACP_procsys_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '后端系统监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(UPACP_procsys_object_list.count(i)))
        row += 1
    for i in set(UPACP_client_id_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '网关监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(UPACP_client_id_object_list.count(i)))
        row += 1
    for i in set(UPACP_Merchant_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '商户监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(UPACP_Merchant_object_list.count(i)))
        row += 1
    for i in set(HTTP_UPACP_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_UPACP_object_list.count(i)))
        row += 1
    for i in set(HTTP_UPACP_qzhlw_object_list):
        ws1.write(row, 0, '全渠道')
        ws1.write(row, 1, 'UPACP')
        ws1.write(row, 2, '接入域监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_UPACP_qzhlw_object_list.count(i)))
        row += 1

    for i in set(QRC_QRC_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'QRC')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QRC_QRC_object_list.count(i)))
        row += 1
    for i in set(QRC_Center_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'QRC')
        ws1.write(row, 2, '交易中心监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QRC_Center_object_list.count(i)))
        row += 1
    for i in set(QRC_Type_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'QRC')
        ws1.write(row, 2, '交易类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QRC_Type_object_list.count(i)))
        row += 1
    for i in set(QRC_Order_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'QRC')
        ws1.write(row, 2, '订单类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QRC_Order_object_list.count(i)))
        row += 1
    for i in set(HTTP_QRC_All_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QRC_All_object_list.count(i)))
        row += 1
    for i in set(HTTP_QRC_zxhlw_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '接入域监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QRC_zxhlw_object_list.count(i)))
        row += 1
    for i in set(HTTP_QRC_error_object_list):
        ws1.write(row, 0, '二维码')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '错误码监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QRC_error_object_list.count(i)))
        row += 1

    for i in set(FS_FS_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'FS')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(FS_FS_object_list.count(i)))
        row += 1
    for i in set(FS_Center_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'FS')
        ws1.write(row, 2, '交易中心监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(FS_Center_object_list.count(i)))
        row += 1
    for i in set(FS_Type_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'FS')
        ws1.write(row, 2, '交易类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(FS_Type_object_list.count(i)))
        row += 1
    for i in set(FS_procSys_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'FS')
        ws1.write(row, 2, '后端系统监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(FS_procSys_object_list.count(i)))
        row += 1
    for i in set(HTTP_fsasAll_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_fsasAll_object_list.count(i)))
        row += 1
    for i in set(HTTP_FS_qzhlw_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '接入域监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_FS_qzhlw_object_list.count(i)))
        row += 1
    for i in set(HTTP_FS_error_object_list):
        ws1.write(row, 0, '客结')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '错误码监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_FS_error_object_list.count(i)))
        row += 1

    for i in set(QTT_QTTB_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'QTT')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QTT_QTTB_object_list.count(i)))
        row += 1
    for i in set(QTT_Center_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'QTT')
        ws1.write(row, 2, '交易中心监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QTT_Center_object_list.count(i)))
        row += 1
    for i in set(QTT_Type_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'QTT')
        ws1.write(row, 2, '交易类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QTT_Type_object_list.count(i)))
        row += 1
    for i in set(HTTP_QTTall_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QTTall_object_list.count(i)))
        row += 1
    for i in set(HTTP_QTTdomain_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '域名监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QTTdomain_object_list.count(i)))
        row += 1
    for i in set(HTTP_QTThlwqz_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '域名监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QTThlwqz_object_list.count(i)))
        row += 1
    for i in set(HTTP_QTThlw_qz_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, 'API监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QTThlw_qz_object_list.count(i)))
        row += 1
    for i in set(HTTP_QTT_error_object_list):
        ws1.write(row, 0, '条码T')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '错误码监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QTT_error_object_list.count(i)))
        row += 1

    for i in set(QAT_QAT_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'QAT')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QAT_QAT_object_list.count(i)))
        row += 1
    for i in set(QAT_Center_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'QAT')
        ws1.write(row, 2, '交易中心监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QAT_Center_object_list.count(i)))
        row += 1
    for i in set(QAT_Type_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'QAT')
        ws1.write(row, 2, '交易类型监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(QAT_Type_object_list.count(i)))
        row += 1
    for i in set(HTTP_QATall_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '整体对象监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QATall_object_list.count(i)))
        row += 1
    for i in set(HTTP_QATdomain_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '接入域监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QATdomain_object_list.count(i)))
        row += 1
    for i in set(HTTP_QATAPI_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, 'API监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QATAPI_object_list.count(i)))
        row += 1
    for i in set(HTTP_QAT_error_object_list):
        ws1.write(row, 0, '条码A')
        ws1.write(row, 1, 'HTTP')
        ws1.write(row, 2, '错误码监控')
        ws1.write(row, 3, i)
        ws1.write(row, 4, str(HTTP_QAT_error_object_list.count(i)))
        row += 1
    wb.save(fileName)


if __name__ == '__main__':
    ExcelName = str(input('Please input the wanted ExcelName: '))
    col_content = readExcel(ExcelName, 12)
    col_event = readExcel(ExcelName, 44)
    coll = extractData(col_content)
    writeExcelAnalyse(col_content, col_event)
    writeExcelRule(coll)
    # print(extractData(col))
    # print(len(extractData(col)))
