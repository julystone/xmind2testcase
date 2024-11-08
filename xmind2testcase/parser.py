#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import itertools
import json
import logging

from icecream import ic

from xmind2testcase.metadata import TestSuite, TestCase, TestStep

config = {'sep': ' ',
          'valid_sep': '&>+/-',
          'precondition_sep': '\n',
          'summary_sep': '\n----\n',
          'ignore_char': '#!！'
          }
"""
tag：            xmind中的标签信息   label
preconditions:  xmind中的备注信息   note
summary:        xmind中的批注信息   comment
两者都要添加到topic里头，而非块上
都可以做到父节点拼接子节点
'3.6.0>'  会以' > '为拼接用例名的 sep
"""


# Todo 为什么summary没有获得父节点的summary

def xmind_to_testsuites(xmind_content_dict):
    """convert xmind file to `xmind2testcase.metadata.TestSuite` list"""
    suites = []

    for sheet in xmind_content_dict:
        logging.debug('start to parse a sheet: %s', sheet['title'])
        root_topic = sheet['topic']
        sub_topics = root_topic.get('topics', [])

        if sub_topics:
            root_topic['topics'] = filter_empty_or_ignore_topic(sub_topics)
        else:
            logging.warning('This is a blank sheet(%s), should have at least 1 sub topic(test suite)', sheet['title'])
            continue
        suite = sheet_to_suite(root_topic)
        # suite.sheet_name = sheet['title']  # root testsuite has a sheet_name attribute
        logging.debug('sheet(%s) parsing complete: %s', sheet['title'], suite.to_dict())
        suites.append(suite)

    return suites


def filter_empty_or_ignore_topic(topics):
    """Filter blank or start with config.ignore_char topic"""
    result = [topic for topic in topics if not (
            topic['title'] is None or
            topic['title'].strip() == '' or
            topic['title'][0] in config['ignore_char'])]

    for topic in result:
        sub_topics = topic.get('topics', [])
        topic['topics'] = filter_empty_or_ignore_topic(sub_topics)

    return result


def filter_empty_or_ignore_element(values):
    """Filter all empty or ignore XMind elements, especially notes、comments、labels element"""
    result = []
    for value in values:
        if isinstance(value, str) and not value.strip() == '' and not value[0] in config['ignore_char']:
            result.append(value.strip())
    return result


def sheet_to_suite(root_topic):
    """convert a xmind sheet to a `TestSuite` instance"""
    suite = TestSuite()
    root_title = root_topic['title']
    separator = root_title[-1]

    if separator in config['valid_sep']:
        logging.debug('find a valid separator for connecting testcase title: %s', separator)
        config['sep'] = separator  # set the separator for the testcase's title
        root_title = root_title[:-1]
    else:
        config['sep'] = ' '

    suite.name = root_title
    suite.details = root_topic['note']
    suite.sub_suites = []

    for suite_dict in root_topic['topics']:
        suite.sub_suites.append(parse_testsuite(suite_dict))

    return suite


def parse_testsuite(suite_dict):
    testsuite = TestSuite()
    testsuite.name = suite_dict['title']
    testsuite.details = suite_dict['note']
    testsuite.testcase_list = []
    logging.debug('start to parse a testsuite: %s', testsuite.name)

    for cases_dict in suite_dict.get('topics', []):
        for case in recurse_parse_testcase(cases_dict):
            testsuite.testcase_list.append(case)  # 此处将解析的测试用例添加到testsuite的testcase_list中

    logging.debug('testsuite(%s) parsing complete: %s', testsuite.name, testsuite.to_dict())
    return testsuite


def recurse_parse_testcase(case_dict, parent=None):
    if is_testcase_topic(case_dict):
        if parm := is_testcase_parmed(case_dict):
            for combie in gen_orth_com(*(parm.values())):
                # 将case_dict里面对应parm的值，替换为combie，生成新的case_dict
                parm_map = {k: v for k, v in zip(parm.keys(), combie)}
                print(parm_map)
                case_dict['title'] = case_dict['title'].format(**parm_map)
                # case_dict['comment'] = json.dumps(parm_map)
                # 然后调用parse_a_testcase生成测试用例
                # case_dict['comment'] = json.dumps(dict(zip(parm.keys(), combie)))
                case = parse_a_testcase(case_dict, parent)
                return case
        case = parse_a_testcase(case_dict, parent)
        yield case
    else:
        if not parent:
            parent = []

        parent.append(case_dict)

        for child_dict in case_dict.get('topics', []):
            for case in recurse_parse_testcase(child_dict, parent):
                yield case

        parent.pop()


def is_testcase_topic(case_dict):
    """A topic with a priority marker, or no subtopic, indicates that it is a testcase"""
    priority = get_priority(case_dict)
    if priority:
        return True

    children = case_dict.get('topics', [])
    if children:
        return False

    return True


def is_testcase_parmed(case_dict):
    summary = case_dict.get('comment', '')
    if summary:
        try:
            parm = json.loads(summary)
        except json.JSONDecodeError:
            return False
        return parm
    return False


def gen_orth_com(*args):
    """
    生成多个输入参数的正交组合
    :param args: 每个参数的可能取值，输入形式为：参数1取值, 参数2取值, ...
    :return: 所有可能的组合
    """
    # 使用 itertools.product 生成笛卡尔积
    print(args)
    combinations = list(itertools.product(*args))
    for idx, combo in enumerate(combinations):
        print(f"组合 {idx + 1}: {combo}")
    return combinations


def parse_a_testcase(case_dict, parent):
    testcase = TestCase()
    topics = parent + [case_dict] if parent else [case_dict]

    summary, parm_dict = gen_testcase_summary(topics)
    testcase.summary = summary if summary else '无'
    testcase.name = gen_testcase_title(topics)
    # expand_summary_parms(testcase.name, parm_dict)

    preconditions = gen_testcase_preconditions(topics)
    testcase.preconditions = preconditions if preconditions else '无'

    testcase.execution_type = get_execution_type(topics)
    testcase.importance = get_priority(case_dict) or 2

    step_dict_list = case_dict.get('topics', [])
    if step_dict_list:
        testcase.steps = parse_test_steps(step_dict_list)
    # print(testcase.steps)

    # the result of the testcase take precedence over the result of the teststep
    testcase.result = get_test_result(case_dict['markers'])

    if testcase.result == 0 and testcase.steps:
        for step in testcase.steps:
            if step.result == 2:
                testcase.result = 2
                break
            if step.result == 3:
                testcase.result = 3
                break

            testcase.result = step.result  # there is no need to judge where test step are ignored

    if testcase.steps:
        for step in testcase.steps:
            step.actions = step.actions
            step.expected_results = step.expected_results

    logging.debug('finds a testcase: %s', testcase.to_dict())
    return testcase


def get_execution_type(topics):
    labels = [topic.get('label', '') for topic in topics]
    labels = filter_empty_or_ignore_element(labels)
    exe_type = 1
    for item in labels[::-1]:
        if item.lower() in ['自动', 'auto', 'automate', 'automation']:
            exe_type = 2
            break
        if item.lower() in ['手动', '手工', 'manual']:
            exe_type = 1
            break
    return labels


def get_priority(case_dict):
    """Get the topic's priority（equivalent to the importance of the testcase)"""
    if isinstance(case_dict['markers'], list):
        for marker in case_dict['markers']:
            if marker.startswith('priority'):
                return int(marker[-1])


def gen_testcase_title(topics):
    """Link all topic's title as testcase title"""
    titles = [topic['title'] for topic in topics]
    titles = filter_empty_or_ignore_element(titles)

    separator = config['sep']
    # if separator != ' ':
    #     separator = f' {separator} '
    # 只保留titles最后两级的标题
    return separator.join(title_trim(titles))


def title_trim(title):
    """trim topic title to last two levels"""
    if len(title) >= 2:
        return title[-2:]
    else:
        return title


def gen_testcase_preconditions(topics):
    try:
        # 过滤空或被忽略的备注信息
        notes = filter_empty_or_ignore_element([topic.get('note', '') for topic in topics])

        # 使用列表推导式生成预条件字符串
        pre_all = [f'{pre_num}. {pre_note}' for pre_num, pre_note in enumerate(notes, 1)]

        # 返回连接后的预条件字符串
        return config['precondition_sep'].join(pre_all) if pre_all else '无'
    except Exception as e:
        logging.error(f'生成预条件时发生错误: {e}')
        return '生成预条件时出错'


def gen_testcase_summary(topics):
    comments = [topic['comment'] for topic in topics]
    comments = filter_empty_or_ignore_element(comments)
    parm_dict = {}
    if comments:
        try:
            parm_dict = json.loads(comments[0])
        except json.JSONDecodeError:
            logging.warning(f'测试用例的注释信息格式错误: {comments[0]}')
    return config['summary_sep'].join(comments), parm_dict


def parse_test_steps(step_dict_list):
    steps = []

    for step_num, step_dict in enumerate(step_dict_list, 1):
        test_step = parse_a_test_step(step_dict)
        test_step.step_number = step_num
        steps.append(test_step)
    return steps


def parse_a_test_step(step_dict):
    test_step = TestStep()
    test_step.actions = step_dict['title']

    expected_topics = step_dict.get('topics', [])
    if expected_topics:  # have expected result
        expected_topic = expected_topics[0]
        test_step.expected_results = expected_topic['title']  # one test step action, one test expected result
        markers = expected_topic['markers']
        test_step.result = get_test_result(markers)
    else:  # only have test step
        markers = step_dict['markers']
        test_step.result = get_test_result(markers)

    logging.debug('finds a teststep: %s', test_step.to_dict())
    return test_step


def get_test_result(markers):
    """test result: non-execution:0, pass:1, failed:2, blocked:3, skipped:4"""
    if isinstance(markers, list):
        if 'symbol-right' in markers or 'c_simbol-right' in markers:
            result = 1
        elif 'symbol-wrong' in markers or 'c_simbol-wrong' in markers:
            result = 2
        elif 'symbol-pause' in markers or 'c_simbol-pause' in markers:
            result = 3
        elif 'symbol-minus' in markers or 'c_simbol-minus' in markers:
            result = 4
        else:
            result = 0
    else:
        result = 0

    return result
