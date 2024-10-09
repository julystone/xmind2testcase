# encoding: utf-8


import json
import xmind
from xmind2testcase.zentao import xmind_to_zentao_csv_file
from xmind2testcase.utils import get_xmind_testcase_list


def main():
    xmind_file = './tempFiles/testcases_bak.xmind'
    print('Start to convert XMind file: %s' % xmind_file)

    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    print('Convert XMind file to zentao csv file successfully: %s' % zentao_csv_file)

    testcases = get_xmind_testcase_list(xmind_file)
    print(testcases)


if __name__ == '__main__':
    main()
