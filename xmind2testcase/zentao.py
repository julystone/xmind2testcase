#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path
from Utils.Excelize import csv_2_excel

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a zentao csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    file_header = ["ID", "用例名称", "所属模块", "标签", "前置条件", "备注", "步骤描述", "预期结果", "编辑模式",
                   "用例等级",
                   "责任人", "用例状态"]
    zentao_testcase_rows = [file_header]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase, file_header)
        zentao_testcase_rows.append(row)

    zentao_file = xmind_file[:-6] + '.csv'
    if os.path.exists(zentao_file):
        os.remove(zentao_file)

    with open(zentao_file, 'w', encoding='utf8', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(zentao_testcase_rows)
        logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', xmind_file, zentao_file)

    return zentao_file


def csv_2_metersphere(csv_file):
    hide_columns = ['A', 'D', 'F', 'I', 'L', 'K']
    column_widths = {
        'B': 30,
        'C': 10,
        'E': 30,
        'G': 50,
        'H': 50,
    }
    style_dict = {
        'font':
            {
                'font_size': 11,
                'font_color': 'FF0000',
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
            },
        'alignment':
            {
                'wrap_text': True,
                'shrink_to_fit': True
            }

    }

    excel_file_name = csv_file[:-4] + '.xlsx'
    csv_2_excel(csv_file, excel_file_name, hide_columns, column_widths, style_dict)


def gen_a_testcase_row(testcase_dict, file_header):
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])

    mapping = {"ID": '',
               "用例名称": testcase_dict['name'],
               "所属模块": gen_case_module(testcase_dict['product'], testcase_dict['suite']),
               "标签": '',
               "前置条件": testcase_dict['preconditions'],
               "备注": testcase_dict['summary'],
               "步骤描述": case_step,
               "预期结果": case_expected_result,
               "编辑模式": 'STEP',
               "用例等级": gen_case_priority(testcase_dict['importance']),
               "责任人": 'july',
               "用例状态": 'Prepare',
               "执行类型": gen_case_execution_type(testcase_dict['execution_type'])
               }
    row = []
    for case_key in file_header:
        try:
            row.append(mapping[case_key])
        except KeyError:
            row.append('未定义此列头内容')
    return row



def gen_case_module(product_name, module_name):
    return f"/{product_name}/{module_name}/"


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += f"{str(step_dict['step_number'])}. {step_dict['actions'].replace('\n', '').strip()}\n"
        case_expected_result += str(step_dict['step_number']) + '. ' + \
                                step_dict['expected_results'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expected_results', '') else ''
    # 去除最后的空行
    case_step = case_step[:-1]
    case_expected_result = case_expected_result[:-1]

    return case_step, case_expected_result


def gen_case_priority(priority):
    mapping = {1: 'P0', 2: 'P1', 3: 'P2', 4: 'P3'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return 'P2'


def gen_case_execution_type(case_type):
    return case_type
    mapping = {1: 'manual', 2: 'automatic'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return 'manual'


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    print('Conver the xmind file to a zentao csv file succssfully: %s', zentao_csv_file)
