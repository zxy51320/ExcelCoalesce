import xlwings as xw
import shutil
import logging

def openXLSX(path_dict, deals, dateList):
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename=f"{path_dict['new_output_path']}/{dateList[1]} output.log", level=logging.DEBUG, format=LOG_FORMAT)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    logging.debug('Create app completed.')

    try:
        collections = {}
        refunds = {}
        total = {}

        with app:
            for deal in deals:
                try:
                    logging.debug(f"Deal: {deal} starting calculate.")
                    input_path = f"{path_dict['input_path']}/{dateList[1]} {deal} Invoicing Report Crefoport.xlsx"
                    with app.books.open(input_path) as wb:
                        logging.debug('Open input file completed.')
                        resDic = calXLSX(wb, deal)
                        logging.debug(f"Deal: {deal} calculated.")
                        collections[deal] = sum(resDic.values())
                except (IOError, KeyError) as e:
                    logging.error(f"Deal: {deal} error - {str(e)}")
                    raise Exception(f"Deal: {deal} error - {str(e)}")

                try:
                    logging.debug(f"Deal: {deal} starting output.")
                    old_file_path = f"{path_dict['old_output_path']}/{dateList[0]} {deal} Output Account Level Data.csv"
                    new_output_path = f"{path_dict['new_output_path']}/{dateList[1]} {deal} Output Account Level Data.csv"
                    shutil.copyfile(old_file_path, new_output_path)
                    logging.debug(f"Deal: {deal} copy output file completed.")
                    with app.books.open(new_output_path) as wb:
                        if writeXLSX(wb, resDic, dateList[1]):
                            wb.save()
                            logging.debug(f"Write {deal} output file completed.")
                except IOError as e:
                    logging.error(f"Deal: {deal} error - {str(e)}")
                    raise Exception(f"Deal: {deal} error - {str(e)}")

                try:
                    logging.debug(f"Deal: {deal} collecting refund.")
                    input_path = f"{path_dict['input_path']}/{deal} repairs.xlsx"
                    with app.books.open(input_path) as wb:
                        logging.debug('Open refund file completed.')
                        refund = getRefund(wb)
                        logging.debug(f"Deal: {deal} refund collected.")
                        refunds[deal] = refund
                except IOError:
                    logging.debug(f"Deal: {deal} refund file not found.")
                    refunds[deal] = 0

            try:
                logging.debug('Starting create deal-level-analysis.')
                old_file_path = f"{path_dict['old_output_path']}/Deal Level Analysis Output Data.csv"
                new_output_path = f"{path_dict['new_output_path']}/Deal Level Analysis Output Data.csv"
                shutil.copyfile(old_file_path, new_output_path)
                logging.debug('Copy deal-level-analysis file completed.')
                with app.books.open(new_output_path) as wb:
                    if Deal_Level_Output(wb, collections, refunds, total, dateList[1]):
                        wb.save()
                        logging.debug('Write deal-level-analysis file completed.')
            except IOError:
                logging.debug('deal-level-analysis file not found.')

            logging.debug('Jobs completed.')
            return True
    except IOError:
        logging.error('Unknown Error.')
        return False


def calXLSX(wb, deal):
    _resDic = {}
    sheet = wb.sheets[0]
    lastRow = sheet.used_range.last_cell.row - 1

    if lastRow == 1 and deal in ('GEMB', 'ECF', 'BNP 1', 'BNP 2'):
        logging.debug(f"{deal} file: No VariableSymbol found.")
    elif lastRow == 1 and deal not in ('GEMB', 'ECF', 'BNP 1', 'BNP 2'):
        logging.error(f"{deal} file: No VariableSymbol found.")
        raise ValueError
    else:
        variableSymbolList = sheet.range(f"B2:B{lastRow}").value
        paymentList = sheet.range(f"D2:D{lastRow}").value

        for variableSymbol, payment in zip(variableSymbolList, paymentList):
            try:
                variableSymbol = str(int(variableSymbol))
            except TypeError:
                logging.error(f"{deal} file, row {variableSymbol} missing value.")
                continue

            _resDic.setdefault(variableSymbol, 0)
            _resDic[variableSymbol] += payment
            _resDic[variableSymbol] = round(_resDic[variableSymbol], 2)

    return _resDic

def writeXLSX(wb, resDic, date):
    logging.debug('Starting writeXLSX.')
    sheet = wb.sheets[0]
    lastRow = sheet.used_range.last_cell.row
    lastCol = sheet.range('A1').end('right').get_address()
    lastTitle = str(sheet.range(lastCol).value)
    sheet.range(lastCol).value = [lastTitle, f"{date.partition('-')[2]}/01/{date.partition('-')[0]}"]
    logging.debug('Insert column completed.')
    lastCol = sheet.range('A1').end('right').get_address()
    variableSymbolList = sheet.range(f"C2:C{lastRow}").value

    for i, variableSymbol in enumerate(variableSymbolList, start=2):
        try:
            curCel = f"{lastCol[0:-1]}{i}"
            variableSymbol = str(int(variableSymbol))
            sheet.range(curCel).value = resDic.get(variableSymbol, 0)
        except TypeError:
            logging.debug(f"Row {i} value illegal.")

    return True

def getRefund(wb):
    sheet = wb.sheets[0]
    lastRow = sheet.used_range.last_cell.row - 1
    refund = sheet.range(f"D{lastRow}").value
    return refund

def Deal_Level_Output(wb, collections, refunds, total, date):
    logging.debug('Starting fill deal-level-analysis file.')
    sheet = wb.sheets[0]
    lastCol = sheet.range('A2').end('right').get_address()
    lastTitle = str(sheet.range(lastCol).value)
    sheet.range(lastCol).value = [lastTitle, f"{date.partition('-')[2]}/01/{date.partition('-')[0]}"]
    logging.debug('Insert column completed.')
    lastCol = sheet.range('A2').end('right').get_address()
    sheet.range(f"{lastCol[0:-1]}22").value = [f"{date.partition('-')[2]}/01/{date.partition('-')[0]}"]
    sheet.range(f"{lastCol[0:-1]}42").value = [f"{date.partition('-')[2]}/01/{date.partition('-')[0]}"]
    dealsList = sheet.range('A3:A19').value

    for i, deal in enumerate(dealsList, start=3):
        curCel = f"{lastCol[0:-1]}{i}"
        total[deal] = collections.get(deal, 0) + refunds.get(deal, 0)
        sheet.range(curCel).value = collections.get(deal, 0)

    logging.debug('Insert collections completed.')

    for i in range(23, 40):
        curCel = f"{lastCol[0:-1]}{i}"
        sheet.range(curCel).value = refunds.get(dealsList[i - 23], 0)

    logging.debug('Insert refunds completed.')

    for i in range(43, 60):
        curCel = f"{lastCol[0:-1]}{i}"
        sheet.range(curCel).value = total[dealsList[i - 43]]

    logging.debug('Insert total completed.')
    return True
