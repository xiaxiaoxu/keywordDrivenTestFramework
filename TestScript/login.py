#encoding=utf-8
#author-夏晓旭
#encoding=utf-8
from ProjectVar.Var import *
from Util.Exel import *
from Action.Action import *

def execute_test_data_file_case(test_data_file_path):
    test_data_excel_file = ParseExcel(test_data_file_path)#生成解析对象
    test_data_excel_file.set_sheet_by_name("Sheet1")
    print test_data_excel_file.get_default_name()
    #定义几个变量
    command_line = ""#存函数和变量的字符串

    action_name = ""#函数名
    locator_method = ""#定位方式
    lcoator_expression = ""#定位表达式
    action_value = ""#对元素要操作的值
    #有0,1,2,3,4个参数，对应5中情况
    #就判断none的格式，
    for id, row in enumerate(test_data_excel_file.get_all_rows()[1:]):
        # print row[action_name_col_no].value,row[locator_method_col_no].value,row[locator_expression_col_no].value,row[action_value_col_no].value
        action_name = row[action_name_col_no].value
        locator_method = row[locator_method_col_no].value
        locator_expression = row[locator_expression_col_no].value
        action_value = row[action_value_col_no].value

        if locator_method is None and locator_expression is None and action_value is None:
            command_line = action_name + "()"
            print "command line:", command_line
        elif locator_method is not None and locator_expression is not None and action_value is None:
            command_line = action_name + "('" + locator_method + "','" + locator_expression + "')"
            print "command line:", command_line
        elif locator_method is None and locator_expression is None and action_value is not None:
            command_line = "%s(u'%s')" % (action_name, action_value)
            print "command line:", command_line
        else:
            command_line = '%s("%s","%s",u"%s")' % (
                action_name, locator_method, locator_expression, action_value)
            print "command line:", command_line

        try:
            time1 = time.time()
            result = eval(command_line)#执行命令
            elapse_time = "%.2f" % (time.time() - time1)
            test_data_excel_file.write_cell_content(id + 2, action_elapse_time_col_no, elapse_time)
            test_data_excel_file.write_cell_content(id + 2, action_result_col_no, u"成功")
            test_data_excel_file.write_cell_content(id + 2, capture_screen_path_col_no, result)#这个可以不写，没啥用
            test_data_excel_file.save_excel_file()
        except Exception, e:
            elapse_time = "%.2f" % (time.time() - time1)
            print traceback.format_exc()#异常信息
            test_data_excel_file.write_cell_content(id + 2, action_elapse_time_col_no, elapse_time)
            test_data_excel_file.write_cell_content(id + 2, action_result_col_no, u"失败")
            test_data_excel_file.write_cell_content(id + 2, action_excetion_info_col_no, traceback.format_exc())
            result = capture_error_screen()
            test_data_excel_file.write_cell_content(id + 2, capture_screen_path_col_no, result)#这个可以不写，没啥用
            test_data_excel_file.save_excel_file()

if __name__=='__main__':
    execute_test_data_file_case(test_data_excel_path)