import xlwings as xw
import shutil 
import logging

def openXLSX(path_dict,deals,dateList):
    try:
        LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
        logging.basicConfig(filename=path_dict['new_output_path']+'/'+dateList[1]+' output.log', level=logging.DEBUG, format=LOG_FORMAT)
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False
        app.screen_updating=False
        logging.debug('Create app completed.')
        collections = {}
        refunds = {}
        total = {}
        for i in deals:
            try:
                #计算各个VariableSymbol的payment
                logging.debug('Deal:'+ i + ' starting calculate.')
                input_path = path_dict['input_path'] + '/' + dateList[1] + ' ' + i + ' Invoicing Report Crefoport.xlsx'
                wb = app.books.open(input_path)#连接excel文件
                logging.debug('Open input file completed.')
                resDic = calXLSX(wb,i)                
                logging.debug('Deal:'+ i + ' calculated.')
                collections[i] = 0
                #保存当前deal的payment总和到字典中
                for key in resDic:
                    collections[i] += resDic[key]
                wb.close()
            except (IOError):
                logging.error('Deal:'+ i + ' input file or VariableSymbol not found.')
                raise Exception
            try:
                #保存单个deal的output
                logging.debug('Deal:'+ i + ' starting output.')
                old_file_path = path_dict['old_output_path'] + '/' + dateList[0] + ' ' + i + ' Output Account Level Data.csv'
                new_output_path = path_dict['new_output_path'] + '/' + dateList[1] + ' ' + i + ' Output Account Level Data.csv'
                shutil.copyfile(old_file_path,new_output_path)
                logging.debug('Deal:' + i + ' copy output file completed.')
                wb = app.books.open(new_output_path)
                if writeXLSX(wb,resDic,dateList[1]):
                    wb.save()
                    logging.debug('Write ' + i + ' output flie completed.')
                wb.close()
            except IOError:
                logging.error('Deal:'+ i + ' old output file not found.')
                raise Exception
            try:
                #计算refund
                logging.debug('Deal:'+ i + ' collecting refund.')
                input_path = path_dict['input_path'] + '/' + i + ' repairs.xlsx'
                wb = app.books.open(input_path)#连接excel文件
                logging.debug('Open refund file completed.')
                refund = getRefund(wb)
                logging.debug('Deal:'+ i + ' refund collected.')
                refunds[i] = refund
                wb.close()
            except IOError:
                logging.debug('Deal:'+ i + ' refund file not found.')
        try:    
            #保存deal_level_analysis
            logging.debug('Starting create deal-level-analysis.')
            old_file_path = path_dict['old_output_path'] + '/' + 'Deal Level Analysis Output Data.csv'
            new_output_path = path_dict['new_output_path'] + '/' + 'Deal Level Analysis Output Data.csv'
            shutil.copyfile(old_file_path,new_output_path)
            logging.debug('Copy deal-level-analysis file completed.')
            wb = app.books.open(new_output_path)
            if Deal_Level_Output(wb,collections,refunds,total,dateList[1]):
                wb.save()
                logging.debug('Write deal-level-analysis file completed.')
            wb.close()
        except IOError:
                logging.debug('deal-level-analysis file not found.')
        app.quit()
        logging.debug('Jobs completed.')
        return True
    except IOError:
        logging.error('Unknown Error.')
        return False

def calXLSX(wb,deal):
    _resDic = {}
    _lastRow = str(wb.sheets[0].used_range.last_cell.row - 1)
    if _lastRow == '1' and deal in ('GEMB','ECF','BNP 1','BNP 2'): 
        logging.debug(deal+' file: No VariableSymbol found.')
    elif _lastRow == '1' and deal not in ('GEMB','ECF','BNP 1','BNP 2'):
        logging.error(deal+' file: No VariableSymbol found.')
        raise ValueError
    else:
        _VariableSymbolList = wb.sheets[0].range('B2:B'+_lastRow).value        
        _PaymentList = wb.sheets[0].range('D2:D'+_lastRow).value
        for i in range (len(_VariableSymbolList)):
            #value方法只能返回浮点数，需要转换成字符串
            try:
                _VariableSymbolList[i] = str(int(_VariableSymbolList[i]))
            except TypeError:
                logging.error(deal+' file, row '+i+' missing value.')
                pass
            try:
                _resDic[_VariableSymbolList[i]] += _PaymentList[i]
                #解决某些浮点数出现长串0的问题
                _resDic[_VariableSymbolList[i]] = round(_resDic[_VariableSymbolList[i]],2)
            except KeyError:
                _resDic[_VariableSymbolList[i]] = _PaymentList[i]
    return _resDic

def writeXLSX(wb,resDic,date):
    logging.debug('Starting writeXLSX.')
    _lastRow = str(wb.sheets[0].used_range.last_cell.row)
    #没发现原生插入列的方法，只能先找到原来的最后一列，再用行插入的方法新增一列
    _lastCol = wb.sheets[0].range('A1').end('right').get_address()
    _lastTitle = str(wb.sheets[0].range(_lastCol).value)
    wb.sheets[0].range(_lastCol).value = [_lastTitle ,date.partition('-')[2] + '/01/' + date.partition('-')[0]] #新增一列
    logging.debug('Insert column completed.')
    _lastCol = wb.sheets[0].range('A1').end('right').get_address() #获取更新后最后一列的地址   
    _VariableSymbolList = wb.sheets[0].range('C2:C'+_lastRow).value
    for i in range (2,len(_VariableSymbolList)+2):
        try:
            _curCel = _lastCol[0:-1] + str(i)
            #value方法只能返回浮点数，需要转换成字符串
            _VariableSymbolList[i-2] = str(int(_VariableSymbolList[i-2]))
            try:
                wb.sheets[0].range(_curCel).value = resDic[_VariableSymbolList[i-2]]
            except KeyError:
                wb.sheets[0].range(_curCel).value = 0
        except TypeError:
            logging.debug('Row '+ str(i) + ' value illegal.')
    return True

def getRefund(wb):
    _lastRow = str(wb.sheets[0].used_range.last_cell.row - 1)
    _refund = wb.sheets[0].range('D'+_lastRow).value
    return _refund

def Deal_Level_Output(wb,collections,refunds,total,date):
    logging.debug('Starting fill deal-level-analysis file.')
    #_lastRow = str(wb.sheets[0].used_range.last_cell.row)
    #没发现原生插入列的方法，只能先找到原来的最后一列，再用行插入的方法新增一列
    _lastCol = wb.sheets[0].range('A2').end('right').get_address()
    _lastTitle = str(wb.sheets[0].range(_lastCol).value)
    wb.sheets[0].range(_lastCol).value = [_lastTitle ,date.partition('-')[2] + '/01/' + date.partition('-')[0]] #新增一列
    logging.debug('Insert column completed.')
    _lastCol = wb.sheets[0].range('A2').end('right').get_address() #获取更新后最后一列的地址
    wb.sheets[0].range(_lastCol[0:-1] + str(22)).value = [date.partition('-')[2] + '/01/' + date.partition('-')[0]]
    wb.sheets[0].range(_lastCol[0:-1] + str(42)).value = [date.partition('-')[2] + '/01/' + date.partition('-')[0]]
    _dealsList = wb.sheets[0].range('A3:A19').value
    for i in range (3,20):
        _curCel = _lastCol[0:-1] + str(i)
        try:
            total[_dealsList[i-3]] = collections[_dealsList[i-3]] + refunds[_dealsList[i-3]]
        except KeyError:
            total[_dealsList[i-3]] = collections[_dealsList[i-3]]
        finally:
            wb.sheets[0].range(_curCel).value = collections[_dealsList[i-3]]
    logging.debug('Insert collections completed.')
    for i in range (23,40):
        _curCel = _lastCol[0:-1] + str(i)
        try:
            wb.sheets[0].range(_curCel).value = refunds[_dealsList[i-23]]
        except KeyError:
            logging.debug('Deal:'+ _dealsList[i-23] + 'does not exist in refunds.')
            wb.sheets[0].range(_curCel).value = 0
    logging.debug('Insert refunds completed.')
    for i in range (43,60):
        _curCel = _lastCol[0:-1] + str(i)
        wb.sheets[0].range(_curCel).value = total[_dealsList[i-43]]
    logging.debug('Insert total completed.')
    return True
