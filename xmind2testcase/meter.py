#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os

from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path


class XMindToZentaoConverter:
    def __init__(self, xmind_file):
        self.xmind_file = get_absolute_path(xmind_file)
        self.zentao_file = self.xmind_file[:-6] + '.csv'
        self.fileheader = ["ID", "用例名称", "所属模块", "标签", "前置条件", "备注", "步骤描述", "预期结果", "编辑模式",
                           "用例等级", "责任人", "用例状态", "excution_type"]
        self.zentao_testcase_rows = [self.fileheader]

    def convert(self):
        """Convert XMind file to a zentao csv file"""
        logging.info('Start converting XMind file(%s) to zentao file...', self.xmind_file)
        testcases = get_xmind_testcase_list(self.xmind_file)

        for testcase in testcases:
            row = self.gen_a_testcase_row(testcase)
            self.zentao_testcase_rows.append(row)

        if os.path.exists(self.zentao_file):
            os.remove(self.zentao_file)

        with open(self.zentao_file, 'w', encoding='utf8', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(self.zentao_testcase_rows)
            logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', self.xmind_file,
                         self.zentao_file)

        return self.zentao_file

    def csv_2_metersphere(self, csv_file):
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
        hide_columns = ['A', 'D', 'F', 'I', 'L', 'K']
        from Utils.Excelize import csv_2_excel
        output_file = csv_file[:-4] + '.xlsx'
        csv_2_excel(csv_file, output_file, hide_columns, column_widths, style_dict)

    def gen_a_testcase_row(self, testcase_dict):
        case_id = ""
        case_title = testcase_dict['name']
        case_module = self.gen_case_module(testcase_dict['product'], testcase_dict['suite'])
        case_tag = ''
        case_preconditions = testcase_dict['preconditions']  # 前置条件
        case_note = testcase_dict['summary']  # 备注
        case_step, case_expected_result = self.gen_case_step_and_expected_result(testcase_dict['steps'])
        case_edit_type = 'STEP'
        case_priority = self.gen_case_priority(testcase_dict['importance'])
        case_owner = 'july'
        case_status = 'Prepare'
        case_execution_type = self.gen_case_execution_type(testcase_dict['execution_type'])

        row = [case_id, case_title, case_module, case_tag, case_preconditions, case_note, case_step,
               case_expected_result, case_edit_type, case_priority, case_owner, case_status, case_execution_type]
        return row

    def gen_case_module(self, product_name, module_name):
        return f"/{product_name}/{module_name}/"

    def gen_case_step_and_expected_result(self, steps):
        case_step = ''
        case_expected_result = ''

        for step_dict in steps:
            case_step += f"{str(step_dict['step_number'])}. {step_dict['actions'].replace('\n', '').strip()}\n"
            case_expected_result += str(step_dict['step_number']) + '. ' + \
                                    step_dict['expected_results'].replace('\n', '').strip() + '\n' \
                if step_dict.get('expected_results', '') else ''
        case_step = case_step[:-1]  # 去除最后的空行
        case_expected_result = case_expected_result[:-1]

        return case_step, case_expected_result

    def gen_case_priority(self, priority):
        mapping = {1: 'P0', 2: 'P1', 3: 'P2', 4: 'P3'}
        return mapping.get(priority, 'P2')

    def gen_case_execution_type(self, case_type):
        mapping = {1: 'manual', 2: 'automatic'}
        return mapping.get(case_type, 'manual')


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    converter = XMindToZentaoConverter(xmind_file)
    zentao_csv_file = converter.convert()
    print('Convert the xmind file to a zentao csv file successfully: %s' % zentao_csv_file)
