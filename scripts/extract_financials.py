"""
IPO & M&A DCF 建模 — PDF 财务数据提取工具

功能：
- 从年报/招股书/审计报告 PDF 中提取财务数据
- 支持合并资产负债表、利润表、现金流量表
- 支持折旧摊销附注提取
- 支持在建工程明细提取
- 支持募投项目参数提取

依赖：pip install PyMuPDF
"""

import fitz  # PyMuPDF
import re
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple


class FinancialExtractor:
    """PDF 财务数据提取器"""

    def __init__(self, pdf_path: str):
        self.pdf_path = Path(pdf_path)
        self.doc = fitz.open(str(self.pdf_path))
        self.full_text = ""
        self.pages_text = []

    def extract_all_text(self):
        """提取全部文本"""
        for i, page in enumerate(self.doc):
            text = page.get_text()
            self.pages_text.append(text)
            self.full_text += text + "\n"
        return self.full_text

    def is_scanned(self) -> bool:
        """检测是否为扫描版 PDF"""
        text_len = len(self.full_text.strip())
        if text_len == 0:
            # 二次确认：检查是否有图片
            for page in self.doc:
                if page.get_images():
                    return True
        return text_len < 100  # 文本极少也判定为扫描版

    def search_keyword(self, keyword: str, context_lines: int = 5) -> List[str]:
        """搜索关键词并返回上下文"""
        results = []
        pattern = re.compile(keyword, re.IGNORECASE)
        lines = self.full_text.split('\n')
        for i, line in enumerate(lines):
            if pattern.search(line):
                start = max(0, i - context_lines)
                end = min(len(lines), i + context_lines + 1)
                context = '\n'.join(lines[start:end])
                results.append(context)
        return results

    def extract_numbers_from_text(self, text: str) -> List[float]:
        """从文本中提取所有数字"""
        # 匹配格式：1,234.56 或 -1,234.56 或 1234.56
        pattern = r'-?[\d,]+\.?\d*'
        matches = re.findall(pattern, text)
        numbers = []
        for m in matches:
            try:
                num = float(m.replace(',', ''))
                numbers.append(num)
            except ValueError:
                continue
        return numbers

    def find_financial_table(self, table_name: str) -> Optional[str]:
        """查找财务报表"""
        # 搜索 "合并资产负债表" / "合并利润表" / "合并现金流量表"
        pattern = re.compile(table_name, re.IGNORECASE)
        for i, page_text in enumerate(self.pages_text):
            if pattern.search(page_text):
                # 收集后续 3 页的文本（报表通常跨多页）
                table_text = page_text
                for j in range(1, 4):
                    if i + j < len(self.pages_text):
                        table_text += "\n" + self.pages_text[i + j]
                return table_text
        return None

    def extract_balance_sheet_items(self) -> Dict[str, float]:
        """提取资产负债表关键项目（万元）"""
        text = self.find_financial_table("合并资产负债表")
        if not text:
            return {}

        items = {}
        patterns = {
            "货币资金": r"货币资金[.\s]*?(-?[\d,]+\.?\d*)",
            "应收账款": r"应收账款[.\s]*?(-?[\d,]+\.?\d*)",
            "应收票据": r"应收票据[.\s]*?(-?[\d,]+\.?\d*)",
            "应收款项融资": r"应收款项融资[.\s]*?(-?[\d,]+\.?\d*)",
            "存货": r"存货[.\s]*?(-?[\d,]+\.?\d*)",
            "应付账款": r"应付账款[.\s]*?(-?[\d,]+\.?\d*)",
            "固定资产": r"固定资产[.\s]*?(-?[\d,]+\.?\d*)",
            "在建工程": r"在建工程[.\s]*?(-?[\d,]+\.?\d*)",
            "短期借款": r"短期借款[.\s]*?(-?[\d,]+\.?\d*)",
            "长期借款": r"长期借款[.\s]*?(-?[\d,]+\.?\d*)",
            "一年内到期": r"一年内到期[的]*非流动负债[.\s]*?(-?[\d,]+\.?\d*)",
            "应付债券": r"应付债券[.\s]*?(-?[\d,]+\.?\d*)",
        }

        for name, pattern in patterns.items():
            match = re.search(pattern, text)
            if match:
                try:
                    value = float(match.group(1).replace(',', ''))
                    items[name] = value / 10000  # 元→万元
                except (ValueError, IndexError):
                    continue

        return items

    def extract_income_statement_items(self) -> Dict[str, float]:
        """提取利润表关键项目（万元）"""
        text = self.find_financial_table("合并利润表")
        if not text:
            return {}

        items = {}
        patterns = {
            "营业收入": r"营业[总]*收入[.\s]*?(-?[\d,]+\.?\d*)",
            "营业成本": r"营业[总]*成本[.\s]*?(-?[\d,]+\.?\d*)",
            "税金及附加": r"税金及附加[.\s]*?(-?[\d,]+\.?\d*)",
            "销售费用": r"销售费用[.\s]*?(-?[\d,]+\.?\d*)",
            "管理费用": r"管理费用[.\s]*?(-?[\d,]+\.?\d*)",
            "研发费用": r"研发费用[.\s]*?(-?[\d,]+\.?\d*)",
            "财务费用": r"财务费用[.\s]*?(-?[\d,]+\.?\d*)",
            "利息费用": r"利息费用[.\s]*?(-?[\d,]+\.?\d*)",
            "利息收入": r"利息收入[.\s]*?(-?[\d,]+\.?\d*)",
            "利润总额": r"利润总额[.\s]*?(-?[\d,]+\.?\d*)",
            "所得税费用": r"所得税费用[.\s]*?(-?[\d,]+\.?\d*)",
            "归[属母]*净利润": r"归属[于]*母[公司]*[股东]*[的]*净利润[.\s]*?(-?[\d,]+\.?\d*)",
        }

        for name, pattern in patterns.items():
            match = re.search(pattern, text)
            if match:
                try:
                    value = float(match.group(1).replace(',', ''))
                    items[name] = value / 10000  # 元→万元
                except (ValueError, IndexError):
                    continue

        return items

    def extract_depreciation_amortization(self) -> Dict[str, float]:
        """提取折旧摊销四项（万元）"""
        # 搜索"现金流量表补充资料" 或 "将净利润调节为经营活动现金流量"
        keyword = "将净利润调节为经营活动现金流量"
        text = ""
        for i, page_text in enumerate(self.pages_text):
            if keyword in page_text:
                # 收集后续页面
                for j in range(4):
                    if i + j < len(self.pages_text):
                        text += self.pages_text[i + j] + "\n"
                break

        if not text:
            return {}

        items = {}
        patterns = {
            "固定资产折旧": r"固定[资产]*折旧[.\s]*?(-?[\d,]+\.?\d*)",
            "无形资产摊销": r"无形资产摊销[.\s]*?(-?[\d,]+\.?\d*)",
            "长期待摊费用摊销": r"长期待摊费用摊销[.\s]*?(-?[\d,]+\.?\d*)",
            "使用权资产折旧": r"使用权资产折旧[.\s]*?(-?[\d,]+\.?\d*)",
        }

        for name, pattern in patterns.items():
            match = re.search(pattern, text)
            if match:
                try:
                    value = float(match.group(1).replace(',', ''))
                    items[name] = value / 10000
                except (ValueError, IndexError):
                    continue

        # 计算合计
        if items:
            items["折旧摊销合计"] = sum(items.values())

        return items

    def extract_capex(self) -> Optional[float]:
        """提取 CAPEX（万元）"""
        text = self.find_financial_table("合并现金流量表")
        if not text:
            return None

        pattern = r"购建固定资产[、]*无形资产和其他长期资产支付[的]*现金[.\s]*?(-?[\d,]+\.?\d*)"
        match = re.search(pattern, text)
        if match:
            try:
                return float(match.group(1).replace(',', '')) / 10000
            except (ValueError, IndexError):
                pass
        return None

    def extract_ipo_fundraising_projects(self) -> List[Dict]:
        """提取募投项目信息（仅适用于招股说明书）"""
        # 搜索 "募集资金运用" / "募集资金投资项目"
        keywords = ["募集资金运用", "募集资金投资项目"]
        text = ""
        for kw in keywords:
            for i, page_text in enumerate(self.pages_text):
                if kw in page_text:
                    for j in range(10):  # 收集后续10页
                        if i + j < len(self.pages_text):
                            text += self.pages_text[i + j] + "\n"
                    break
            if text:
                break

        if not text:
            return []

        projects = []
        # 尝试提取项目名称和总投资
        # 格式：项目名称 ... 总投资 XX万元 ... 募集资金投入 XX万元
        project_pattern = re.compile(
            r'([\u4e00-\u9fa5A-Za-z0-9（）()\-、，,.]+?项目)[\s\S]*?'
            r'(?:总投资|投资总额)[^\d]*?([\d,]+\.?\d*)[\s\S]*?'
            r'(?:募集资金|拟[用使]募集)[^\d]*?([\d,]+\.?\d*)',
            re.IGNORECASE
        )

        for match in project_pattern.finditer(text):
            try:
                name = match.group(1).strip()
                total_invest = float(match.group(2).replace(',', '')) / 10000
                fund_invest = float(match.group(3).replace(',', '')) / 10000
                projects.append({
                    "项目名称": name,
                    "总投资(万元)": total_invest,
                    "募集资金投入(万元)": fund_invest
                })
            except (ValueError, IndexError):
                continue

        return projects

    def extract_report_date(self) -> Optional[str]:
        """提取报告期截止日"""
        # 搜索日期格式
        patterns = [
            r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日',
            r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
        ]

        # 优先从封面/首页搜索
        for i in range(min(5, len(self.pages_text))):
            for pattern in patterns:
                match = re.search(pattern, self.pages_text[i])
                if match:
                    year, month, day = match.groups()
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"

        return None

    def close(self):
        """关闭 PDF 文档"""
        self.doc.close()

    def to_json(self, output_path: str):
        """将提取结果保存为 JSON"""
        data = {
            "文件": str(self.pdf_path),
            "是否扫描版": self.is_scanned(),
            "报告期截止日": self.extract_report_date(),
            "资产负债表": self.extract_balance_sheet_items(),
            "利润表": self.extract_income_statement_items(),
            "折旧摊销": self.extract_depreciation_amortization(),
            "CAPEX": self.extract_capex(),
            "募投项目": self.extract_ipo_fundraising_projects(),
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return data


# ============================================================
# 命令行入口
# ============================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("用法: python extract_financials.py <PDF文件路径> [输出JSON路径]")
        print("示例: python extract_financials.py 招股书.pdf output.json")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "extracted_data.json"

    print(f"正在提取 {pdf_path} ...")
    extractor = FinancialExtractor(pdf_path)
    extractor.extract_all_text()

    if extractor.is_scanned():
        print("⚠ 检测到扫描版 PDF，部分数据可能无法提取。")

    data = extractor.to_json(output_path)
    extractor.close()

    print(f"✅ 数据已保存到 {output_path}")
    print(f"   - 资产负债表项目数: {len(data.get('资产负债表', {}))}")
    print(f"   - 利润表项目数: {len(data.get('利润表', {}))}")
    print(f"   - 折旧摊销项目数: {len(data.get('折旧摊销', {}))}")
    print(f"   - 募投项目数: {len(data.get('募投项目', []))}")
