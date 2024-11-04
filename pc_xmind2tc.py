# encoding: utf-8


import json
import xmind
from icecream import ic

from xmind2testcase.zentao import xmind_to_zentao_csv_file, csv_2_metersphere
from xmind2testcase.utils import get_xmind_testcase_list


# TODO  Xmind8中的  summary怎么用上 - finish
# TODO  Xmind8中的  link怎么用上
# TODO  Xmind8，需要重复创建相似用例，怎么使用已有字段
# TODO  Excelize.py文件函数整理 - finish
# TODO  可支持字段的config配置项抽离

def main():
    xmind_file = './tempFiles/testcases_template.xmind'
    print('Start to convert XMind file: %s' % xmind_file)

    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    csv_2_metersphere(zentao_csv_file)
    print('Convert XMind file to zentao csv file successfully: %s' % zentao_csv_file)

    testcases = get_xmind_testcase_list(xmind_file)
    ic(testcases)


if __name__ == '__main__':
    main()