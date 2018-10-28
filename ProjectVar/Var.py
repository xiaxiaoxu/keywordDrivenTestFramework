#encoding=utf-8
#author-夏晓旭
#encoding=utf-8
import os

#获取工程所在的目录的绝对路径
project_path=os.path.dirname(os.path.dirname(__file__))

#获取页面对象库文件的绝对路径
#page_object_repository_path=project_path.decode("utf-8")+u"/Conf/PageObjectRepository.ini"

#测试数据excel文件的绝对路径
test_data_excel_path=project_path.decode("utf-8")+u"/TestData/126邮箱的测试用例.xlsx"

#浏览器驱动文件所在的绝对路径
ieDriverFilePath="d:\\IEDriverServer"
chromeDriverFilePath="c:\\chromedriver"
firefoxDriverFilePath="c:\\geckodriver"


#excel列的含义
action_name_col_no=2
locator_method_col_no=3
locator_expression_col_no=4
action_value_col_no=5
action_elapse_time_col_no=7
action_result_col_no=8
action_excetion_info_col_no=9
capture_screen_path_col_no=10