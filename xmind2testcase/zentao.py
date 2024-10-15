#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a zentao csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["ID", "用例名称", "所属模块", "标签", "前置条件", "备注", "步骤描述", "预期结果", "编辑模式",
                  "用例等级",
                  "责任人", "用例状态"]
    zentao_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        zentao_testcase_rows.append(row)

    zentao_file = xmind_file[:-6] + '.csv'
    if os.path.exists(zentao_file):
        os.remove(zentao_file)
        # logging.info('The zentao csv file already exists, return it directly: %s', zentao_file)
        # return zentao_file

    with open(zentao_file, 'w', encoding='utf8', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(zentao_testcase_rows)
        logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', xmind_file, zentao_file)

    return zentao_file


def csv_2_metersphere(csv_file):
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
    from Utils.Excelize import csv_2_excel, ReadExcel
    output_file = csv_file[:-4] + '.xlsx'
    out = csv_2_excel(csv_file, output_file, column_widths, style_dict)
    po = ReadExcel(out)
    po.hide_column(['A', 'D', 'F', 'I', 'L', 'K'])
    po.save()


def gen_a_testcase_row(testcase_dict):
    """
    {
        "name": "设置相关 - 默认值：旧版",
        "product": "3.7.0",
        "suite": "证券风格"
        "version": 1,
        "summary": "外层备注\n----\n内层备注",
        "preconditions": "外层前置条件\n----\n内层前置条件",
        "execution_type": 1,
        "importance": 2,
        "estimated_exec_duration": 3,
        "status": 7,
        "result": 0,
        "steps": [
            {
                "step_number": 1,
                "actions": "进入设置界面",
                "expectedresults": "",
                "execution_type": 1,
                "result": 0
            },
            {
                "step_number": 2,
                "actions": "观察证券风格设置开关",
                "expectedresults": "默认旧版风格",
                "execution_type": 1,
                "result": 0
            }
        ],
    },
    """
    case_id = ""
    case_title = testcase_dict['name']
    case_module = gen_case_module(testcase_dict['product'], testcase_dict['suite'])
    case_tag = ''
    case_preconditions = testcase_dict['preconditions']  # 前置条件
    case_note = testcase_dict['summary']  # 备注
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    case_edit_type = 'STEP'
    case_priority = gen_case_priority(testcase_dict['importance'])
    case_owner = 'july'
    case_status = 'Prepare'
    case_execution_type = gen_case_execution_type(testcase_dict['execution_type'])
    row = [case_id, case_title, case_module, case_tag, case_preconditions, case_note, case_step, case_expected_result,
           case_edit_type, case_priority, case_owner, case_status]
    return row


def gen_case_module(product_name, module_name):
    return f"/{product_name}/{module_name}/"


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result += str(step_dict['step_number']) + '. ' + \
                                step_dict['expectedresults'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''
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
    mapping = {1: 'manual', 2: 'automatic'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return 'manual'


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    print('Conver the xmind file to a zentao csv file succssfully: %s', zentao_csv_file)
