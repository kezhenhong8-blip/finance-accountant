# -*- coding: utf-8 -*-
"""发票票种识别工具（skill: financial-check 配套）

对每张发票 PDF 自动识别：
  - 票种：增值税专用发票 / 增值税普通发票 / 数电票（专/普）/ 红字发票 / 铁路电子客票 / 机动车票 / 其他
  - 关键字段：购买方/销售方名称、价税合计（小写）、金额与税额、发票号码、票价（铁路客票）
  - 输出：逐文件结果 + 依据描述（文本标题/版面特征/文件名启发式）

用法（PowerShell）:
    $env:PYTHONIOENCODING="utf-8"
    python invoice_detect.py "D:/.../支出发票/xxx.pdf"
    python invoice_detect.py "." --dir "D:/.../支出发票"    # 批量扫描（递归 **/*.pdf）

退出码: 0 = 全部确定；1 = 存在未确定票种（须人工核对，不得按默认票种处理）
依赖: pdfplumber（可选；缺失时仅做文件名启发式）
规则详见 docs/invoice-recognition.md。
"""
import argparse
import glob
import os
import re
import sys

TITLE_PATTERNS = [
    ('增值税专用发票', re.compile(r'增值税专用发票|增值税电子专用发票'), 10),
    ('增值税普通发票', re.compile(r'增值税普通发票|增值税电子普通发票'), 10),
    ('数电票(专)', re.compile(r'电子发票\s*[（(]\s*(增值税)?专用发票'), 9),
    ('数电票(普)', re.compile(r'电子发票\s*[（(]\s*(增值税)?普通发票'), 9),
    ('铁路电子客票', re.compile(r'铁路电子客票'), 9),
    ('机动车销售统一发票', re.compile(r'机动车销售统一发票'), 9),
    ('二手车销售统一发票', re.compile(r'二手车销售统一发票'), 9),
    ('支付/电子回单(非发票)', re.compile(r'电子回单|交易回单|转账回单|行程校验单|航旅纵横'), 8),
]
RED_TRAIL = re.compile(r'红字发票|红\s*冲|红冲|作废数电|负数')

def extract_text(path):
    try:
        import pdfplumber
    except ImportError:
        return None, 'no_pdfplumber'
    try:
        with pdfplumber.open(path) as pdf:
            text = '\n'.join((p.extract_text() or '') for p in pdf.pages)
        return text, 'ok'
    except Exception as e:
        return None, 'error:' + str(e)

def classify(text, filename):
    reasons = []
    best = ('其他', 0)
    if text:
        for name, pat, score in TITLE_PATTERNS:
            if pat.search(text):
                reasons.append('文本含"%s"' % name)
                if score > best[1]:
                    best = (name, score)
        if RED_TRAIL.search(text):
            best = ('红字发票相关', 12)
            reasons.append('含红冲/红字/负数票价等字样(优先)')
    if best[1] == 0:
        if filename and re.search(r'回单', filename):
            best = ('支付/电子回单(非发票-文件名)', 7)
            reasons.append('文件名含"回单"')
        elif filename and re.search(r'红', filename):
            best = ('疑似红字(文件名)', 3)
            reasons.append('文件名含"红"')
        elif filename and re.search(r'专票|专用', filename):
            best = ('疑似专票(文件名)', 3)
            reasons.append('文件名含"专票/专用"')
        elif filename and re.search(r'普票|普通', filename):
            best = ('疑似普票(文件名)', 3)
            reasons.append('文件名含"普票/普通"')
        elif text:
            seg = text.split('销售方')[0].split('购买方')[-1]
            if re.search(r'开户行|银行账号|开户银行', seg):
                best = ('疑似专票(版面:购买方含开户行)', 2)
                reasons.append('购买方信息段含开户行/账号(弱证据)')
    return best[0], reasons

def extract_fields(text):
    fields = {}
    if not text:
        return fields
    m = re.search(r'买\s*方?\s*名称[：:]\s*([^\n]*)', text) or \
        re.search(r'购买方[^\n:：]{0,6}[：:]?\s*名称[：:]\s*([^\n]*)', text)
    if m:
        fields['购买方'] = m.group(1).strip()
    m = re.search(r'售\s*方?\s*名称[：:]\s*([^\n]*)', text) or \
        re.search(r'销售方[^\n:：]{0,6}[：:]?\s*名称[：:]\s*([^\n]*)', text)
    if m:
        fields['销售方'] = m.group(1).strip()
    m = re.search(r'[（(]小写[）)]\s*[¥￥]?\s*([\d,]+\.\d{2})', text)
    if m:
        fields['价税合计'] = m.group(1)
    else:
        m = re.search(r'价税合计[^¥￥\n]{0,15}[¥￥]?\s*([\d,]+\.\d{2})', text)
        if m:
            fields['价税合计'] = m.group(1)
    m = re.search(r'合\s*计\s*[¥￥]\s*([\d,]+\.\d{2})\s*[¥￥]?\s*([\d,]+\.\d{2})?', text)
    if m:
        fields['合计金额'] = m.group(1)
        if m.group(2):
            fields['合计税额'] = m.group(2)
    m = re.search(r'票价[：:]\s*[¥￥]?\s*(-?[\d,]+\.\d{2})', text)
    if m:
        fields['票价'] = m.group(1)
    m = re.search(r'发票号码[：:]\s*([\d-]{8,30})', text)
    if m:
        fields['发票号码'] = m.group(1).replace(' ', '')
    return fields

def main():
    ap = argparse.ArgumentParser(description='发票票种识别')
    ap.add_argument('paths', nargs='+', help='发票PDF路径')
    ap.add_argument('--dir', default=None, help='批量扫描目录（递归 **/*.pdf）')
    args = ap.parse_args()

    targets = list(args.paths)
    if args.dir:
        targets = sorted(glob.glob(os.path.join(args.dir, '**', '*.pdf'), recursive=True))

    if not targets:
        raise SystemExit('未找到 PDF 文件')

    pdfplumber_ok = True
    unknown = []
    total = len(targets)
    done = 0
    for path in targets:
        text, status = extract_text(path)
        if status == 'no_pdfplumber':
            pdfplumber_ok = False
        fname = os.path.basename(path)
        print('文件: ' + fname)
        if text is None:
            print('  -> 文本提取失败(' + status + ')，走文件名启发式')
            text = ''
        kind, reasons = classify(text, fname)
        fields = extract_fields(text) if text else {}
        print('  票种: ' + kind)
        for r in reasons[:4]:
            print('  依据: ' + r)
        for k, v in fields.items():
            print('  ' + k + ': ' + v)
        if kind == '其他' or kind.startswith('疑似'):
            unknown.append(fname)
        else:
            done += 1
        print()

    if not pdfplumber_ok:
        print('提示: 未安装 pdfplumber，仅做了文件名启发式；请 pip install pdfplumber 后重跑以获得准确识别')
    if unknown:
        print('未确定票种 ' + str(len(unknown)) + ' 个 / 共 ' + str(total) +
              '（已确定 ' + str(done) + '），以下请人工核对或查验平台确认（铁律：不得按默认票种处理）:')
        for f in unknown:
            print('  ' + f)
        sys.exit(1)
    else:
        print('全部 ' + str(done) + ' 个票种识别完成')
        sys.exit(0)

if __name__ == '__main__':
    main()