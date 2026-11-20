from pathlib import Path
import json
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts' / '清洗账表.py'


class 清洗行为(unittest.TestCase):
    def setUp(self):
        self.work = Path(tempfile.mkdtemp(prefix='账表清洗验证-'))
        self.journal = self.work / '任意名称甲.xlsx'
        self.balance = self.work / '任意名称乙.xlsx'
        self.output = self.work / '结果'
        self.make_journal()
        self.make_balance()

    def make_journal(self, duplicate=False, loan_conflict=False):
        book = Workbook()
        book.active.title = '封面'
        sheet = book.create_sheet('任意凭证页')
        sheet.merge_cells('A2:H2')
        sheet['A2'] = '导出记录'
        for region, value in [('A6:B6','本位币'),('C6:C7','会计日期'),('D6:D7','凭证编号'),
                              ('E6:E7','账户组合编码'),('F6:F7','账户说明'),('G6:G7','摘要'),
                              ('H6:H7','科目名称'),('I6:I7','行号')]:
            sheet.merge_cells(region)
            sheet[region.split(':')[0]] = value
        sheet['A7'], sheet['B7'] = '贷方金额', '借方金额'
        sheet['J6'], sheet['J7'] = '原币', '借方金额'
        if duplicate:
            sheet['K6'], sheet['K7'] = '本位币', '借方金额'
        sheet.append((0,100,'2025/01/02','000001','001.0.0.2501010107.0.0.009.S03.B01.0',
                      '测试公司.缺省.-.长期借款-人民币-本金-质押加抵押加保证借款.-.缺省.合同利率5.25%.示例银行城东支行.示例银行.缺省',
                      '本金','长期借款-人民币-本金-质押加抵押加保证借款',1,9))
        sheet.append((100,0,'2025/01/02','000001','001.0.0.2231020101.0.0.009.001.0.0',
                      '测试公司.缺省.-.应付利息-长期借款利息-正常利息.-.缺省.合同利率5.25%.测试公司.缺省.缺省',
                      '利息','应付利息-长期借款利息-正常利息',2,9))
        if loan_conflict:
            sheet.append((0,0,'2025/01/02','000001','001.0.0.2501010107.0.0.009.S04.B04.0',
                          '测试公司.缺省.-.长期借款-人民币-本金-质押加抵押加保证借款.-.缺省.合同利率5.25%.另一银行城西支行.另一银行.缺省',
                          '另一个银行','长期借款-人民币-本金-质押加抵押加保证借款',3,9))
        sheet.append((0,-10,'2025年01月02日','000002','001.0.0.1002010101.00005.0.0.001.B01.0',
                      '测试公司.缺省.-.银行存款-人民币.示例银行城东支行00001234.缺省.-.测试公司.示例银行.缺省',
                      '红字','银行存款-人民币',4,99))
        sheet.append((-10,0,'2025.01.02','000002','001.0.0.2203010101.0.001.009.0.0.0',
                      '测试公司.缺省.-.预收账款-经营性-房产预收款-预收房款.-.销售-住宅.示例项目.缺省.缺省.缺省',
                      '客户缺失','预收账款-经营性-房产预收款-预收房款',5,99))
        book.save(self.journal)

    def make_balance(self, no_cache=False):
        book = Workbook()
        sheet = book.active
        sheet.title = '科目明细'
        sheet['F1'] = '2025年01月至12月'
        for region, value in [('A3:A4','帐户说明'),('B3:B4','账户编码组合'),('C3:D3','本年发生额'),
                              ('E3:E4','期末余额'),('F3:F4','期初余额'),('G3:G4','会计科目编码')]:
            sheet.merge_cells(region)
            sheet[region.split(':')[0]] = value
        sheet['C4'], sheet['D4'] = '贷方', '借方'
        sheet.append(('测试公司-.-.银行存款-人民币.示例银行城东支行00001234-.-.测试公司.示例银行-',
                      '001.0.0.1002010101.00005.0.0.001.B01.0',0,-10,90,100,'1002010101'))
        sheet.append(('测试公司-.-.银行存款-人民币.示例银行城东支行00001234-.-.第二公司.示例银行-',
                      '001.0.0.1002010101.00005.0.0.002.B01.0',0,0,-40,-40,'1002010101'))
        sheet.append(('测试公司-.-.应付利息-长期借款利息-正常利息.--.合同利率5\\.25%.测试公司--',
                      '001.0.0.2231020101.0.0.009.001.0.0',100,0,-100,0,'2231020101'))
        sheet.append(('测试公司-.-.长期借款-人民币-本金-质押加抵押加保证借款.--.合同利率5\\.25%.示例银行城东支行.示例银行-',
                      '001.0.0.2501010107.0.0.009.S03.B01.0',0,100,100,0,'2501010107'))
        if no_cache:
            sheet['D5'] = '=1+2'
        book.save(self.balance)

    def run_skill(self, success=True, files=None, optional=()):
        files = files or [self.balance,self.journal]
        result = subprocess.run([sys.executable, str(SCRIPT), '--输入', *map(str,files),
                                 '--输出目录', str(self.output), *optional],
                                capture_output=True, encoding='utf-8')
        if success:
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            return next(self.output.glob('账表清洗-*'))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(list(self.output.glob('账表清洗-*')))
        return result.stderr + result.stdout

    @staticmethod
    def load_rows(path, name):
        book = load_workbook(path, read_only=True, data_only=True)
        rows = list(book[name].values)
        book.close()
        return rows

    def test_合并表头重排负数和贷款匹配(self):
        result = self.run_skill()
        journal = self.load_rows(result / '序时账-借贷两列.xlsx', '序时账')
        balance = self.load_rows(result / '余额表-借贷两列.xlsx', '余额表')
        self.assertEqual(len(journal), 5)
        self.assertEqual([r[8:10] for r in journal[1:3]], [('S03','示例银行城东支行')]*2)
        self.assertEqual(journal[1][2], '000001')
        self.assertEqual(journal[3][6:10], (-10,0,'00005','示例银行城东支行00001234'))
        self.assertEqual(journal[4][9], '往来未记录')
        self.assertEqual(len(balance), 4)
        self.assertEqual(balance[1][2:4], ('00005','示例银行城东支行00001234'))
        self.assertEqual(balance[1][5:11], (60,None,-10,0,50,None))
        self.assertEqual(balance[2][2:4], ('S03','示例银行城东支行'))
        self.assertTrue((result / '清洗核对.xlsx').is_file())

    def test_重复本位币列报位置(self):
        self.make_journal(duplicate=True)
        error = self.run_skill(success=False)
        self.assertIn('借方', error)
        self.assertIn('K', error)
        choice = self.work / '确认列.json'
        choice.write_text(json.dumps({'序时账':{'借方':'B'}},ensure_ascii=False),encoding='utf-8')
        result = self.run_skill(optional=('--指定列',str(choice)))
        journal = self.load_rows(result/'序时账-借贷两列.xlsx','序时账')
        self.assertEqual(journal[1][6:8],(100,0))

    def test_公式未重算不按零继续(self):
        self.make_balance(no_cache=True)
        error = self.run_skill(success=False)
        self.assertIn('D5', error)
        self.assertIn('公式', error)

    def test_同一合同出现多个贷款行保持待核(self):
        self.make_journal(loan_conflict=True)
        result = self.run_skill()
        journal = self.load_rows(result / '序时账-借贷两列.xlsx', '序时账')
        self.assertEqual(journal[2][9], '贷款银行待核')

    def test_单表自动判断和未知金额单位保护(self):
        result = self.run_skill(files=[self.balance])
        self.assertTrue((result/'余额表-借贷两列.xlsx').is_file())
        self.assertFalse((result/'序时账-借贷两列.xlsx').exists())
        self.output = self.work/'单位结果'
        book = load_workbook(self.balance)
        book.active['A1'] = '金额单位：万元'
        book.save(self.balance)
        self.assertIn('换算',self.run_skill(success=False,files=[self.balance]))


if __name__ == '__main__':
    unittest.main()
